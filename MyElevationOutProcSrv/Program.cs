using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyElevationOutProcSrv
{
    /// <summary>
    /// 独自のアプリケーションコンテキストを作成してアプリケーションの寿命を管理する。
    /// ここでCOMのクラスファクトリをシステムに登録し、クライアントからのCOM生成要求の待ち受けを行う。
    /// 生きているオブジェクト数が0になって一定時間経過したらクラスファクトリの登録を解除し、
    /// アプリケーションを終了させる。
    /// </summary>
    public class MyApplicationContext : ApplicationContext
    {
        #region P/Invoke
        /// <summary>
        /// 登録した全てのクラスファクトリを一斉に受け入れ可能にする
        /// </summary>
        [DllImport("ole32.dll", PreserveSig = false)] // HRESULTの戻り値は例外として受けとる
        static extern void CoResumeClassObjects();

        /// <summary>
        /// 登録した全てのクラスファクトリの受付を一斉に停止する
        /// </summary>
        [DllImport("ole32.dll", PreserveSig = false)] // HRESULTの戻り値は例外として受けとる
        static extern void CoSuspendClassObjects();
        #endregion

        /// <summary>
        /// MyElevationOutProcSrvFactoryクラスファクトリ
        /// </summary>
        private readonly MyElevationOutProcSrvFactory classFactory = new MyElevationOutProcSrvFactory();

        /// <summary>
        /// 現在のコンテキスト
        /// </summary>
        public static MyApplicationContext Current { get; private set; }

        /// <summary>
        /// GCとアイドル経過時間を計測するためのタイマー
        /// </summary>
        private System.Windows.Forms.Timer gcTimer;

        /// <summary>
        /// アクティブなオブジェクト数のカウンタ
        /// </summary>
        private int activeCount;

        /// <summary>
        /// アクティブなオブジェクト数の最大値
        /// </summary>
        private int highWaterMark;

        /// <summary>
        /// アクティブなオブジェクトが0になってから
        /// サーバーを終了するまでの待機時間
        /// </summary>
        public int ShutdownDelaySecs { set; get; } = 3;

        /// <summary>
        /// 最後にアクティブが確認された時刻
        /// </summary>
        private DateTime LastUseTime = DateTime.Now;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MyApplicationContext()
        {
            // このコンテキストを現在のコンテキストとして保存する
            Current = this;

            // アプリケーション終了時
            Application.ApplicationExit += OnApplicationExit;

            // シャットダウン、ログオフが要求された場合のイベント
            SystemEvents.SessionEnding += (object sender, SessionEndingEventArgs e) =>
            {
                Shutdown();
                Application.Exit();
            };

            // クラスファクトリの登録と受付開始
            Start();

            // アクティブなCOM残存数が0になってから一定時間経過したら終了させるためのUIタイマー
            gcTimer = new System.Windows.Forms.Timer();
            gcTimer.Interval = 500; // 500mSec毎
            gcTimer.Tick += OnIntervalTimer;
            gcTimer.Start();
        }

        /// <summary>
        /// COMサーバーのファクトリを登録し、受付を開始する
        /// </summary>
        private void Start()
        {
            Console.WriteLine("Start");

            // クラスファクトリの登録
            classFactory.CoRegisterClassObject();

            // すべてのクラスファクトリで一斉にオブジェクト生成受付を開始する
            CoResumeClassObjects();
        }

        /// <summary>
        /// COMサーバーのファクトリを停止し、登録解除する。
        /// </summary>
        /// <returns>残存オブジェクトが0であるか？</returns>
        private bool Shutdown()
        {
            Console.WriteLine("Shutdown");

            // すべてのクラスファクトリを一斉停止する
            // (この時点で新しいオブジェクトは生成されなくなる)
            CoSuspendClassObjects();

            //　クラスファクトリの登録解除
            classFactory.CoRevokeClassObject();

            return activeCount == 0;
        }

        /// <summary>
        /// 定期的に呼び出されるGUIタイマー。
        /// COMオブジェクトの使用中の残存カウントが0になって一定時間経過したらEXEサーバーを終了せさる。
        /// </summary>
        /// <param name="state"></param>
        private void OnIntervalTimer(object Sender, EventArgs e)
        {
            // 参照されているCOMオブジェクトがReleaseされてもGCが走るまで
            // デストラクタは動かないため、定期的に明示的にGCを行う。
            GC.Collect();

            if (activeCount > 0)
            {
                // まだ生きているオブジェクトがある
                LastUseTime = DateTime.Now;
            }
            else if (highWaterMark > 0)
            {
                // 過去に1つ以上のオブジェクトを生成済みであり、
                // 且つ、残存数が0になってから所定時間を経過した場合は
                // アプリケーションを終了する。
                TimeSpan span = DateTime.Now - LastUseTime;
                if (span.TotalSeconds > ShutdownDelaySecs)
                {
                    if (Shutdown())
                    {
                        Application.Exit();
                    }
                }
            }
        }

        /// <summary>
        /// アプリケーション終了時のイベントハンドラ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnApplicationExit(object sender, EventArgs e)
        {
            Console.WriteLine("OnApplicationExit");
        }

        /// <summary>
        /// アクティブオブジェクト数のカウントアップ
        /// </summary>
        public void IncrementCount()
        {
            int cnt = Interlocked.Increment(ref activeCount);
            if (cnt > highWaterMark)
            {
                highWaterMark = cnt;
            }
            Console.WriteLine("incl count={0} waterMark={0}", activeCount, highWaterMark);
        }

        /// <summary>
        /// アクティブオブジェクト数のカウントダウン
        /// </summary>
        public void DecrementCount()
        {
            Interlocked.Decrement(ref activeCount);
            Console.WriteLine("decl count={0}", activeCount);
        }
    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // 独自のアプリケーションコンテキストを作成してアプリケーションの寿命を管理する
            var appContext = new MyApplicationContext();

            // 設定値(app.config)の取り込み
            appContext.ShutdownDelaySecs = int.Parse(ConfigurationManager.AppSettings["ShutdownDelaySecs"]);

            // アプリケーションのメッセージループを開始する。
            Application.Run(appContext);

            Console.WriteLine("done.");
        }
    }
}
