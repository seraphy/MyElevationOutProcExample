C#でUACの昇格可能なEXEのCOMオブジェクトを作成する手順

# 概要

本サンプルはC#によるOut-of-processのCOMサーバーの作成と、それをUACの昇格可能なCOMにする手順を示すものである。

## COM Elevation Monikerについて

Vista以降、管理者権限が必要な場合はUACによる昇格が必要となっているが、
昇格する単位はプロセス単位であり、且つ、プロセスが起動するときに昇格しなければならない。(起動してから昇格することはできない。)

このため、通常権限で起動したアプリケーション内から管理者権限が必要な処理を行うためには、

- 別の昇格可能なEXEプロセスをShellExecute等で呼び出す。
- もしくは、[COM Elevation Moniker](https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms679687.aspx)を使う

の、いずれかの方法をとることになる。


```ShellExecute``` を使う場合には親プロセスから引数としてパラメータを渡す以上のことをやろうと思うと、プロセス間通信などめんどくさい仕組みが必要になるが、
COMであれば、COMのインフラストラクチャによって、プロセス間通信が単なるメソッドやプロパティのアクセスという平易な形で実現できる。

したがって、プログラムの中から管理者権限のあれこれをやりたいのであれば、昇格可能なCOMで実装したほうが、いろいろ簡単になるであろう。


この昇格可能COMオブジェクトは、```COM Elevation Moniker``` という、COMの複合モニカの仕組みを使って ```"Elevation:administrator!new:{CLSID}"``` のような文字列を指定することで、
**Out-Of-ProcessのCOM**(つまり、EXEのCOM)をUACで昇格して起動する。

オブジエクトの生成時にモニカによってUAC昇格のための処理が挟み込まれるような感じとなる。

- なお、原理的にIn-ProcのCOMでは昇格できない。
  - DLLのCOMを作成した場合でも、明示的に ```CLSCTX_LOCAL_SERVER``` として ```dllhost.exe```プロセス経由でEXEでサロゲートされるように起動すれば昇格可能ではある。(C++からの利用などはCLSCTXを明示できるため。)
  - ただし、WSHやVBAといったCOMクライアントから明示的にCLSCTX_LOCAL_SERVERを指定する方法がなく、その場合は既定でIn-Procで起動されてしまうため、WSHやVBAからの利用があるならばDLLのCOMは適していない。
    - (たとえば、x86のWSHからx64のdllを呼び出す、もしくは、x64のWSHからx86のdllを呼び出す場合には、In-Procでは成立しないのでLocalServerが試行され、結果的に、偶然、うまくゆく場合もありえる。)


昇格可能なCOMオブジェクトを作成するには比較的簡単で、COMのCLSIDのレジストリエントリに

- ```LocalizedString``` という文字列リソースを示す値
- ```Elevation``` キー
  - ```Enabled``` = (DWORD) ```1```

の2つがあれば良い。(それ以外には何も必要ない。他にもアイコンのオプションなどがあるが、必須ではない。)


```LocalizedString``` はUACの昇格ダイアログで表示されるコンポーネント名をリソースから取得するためのリソースキーである。


## C#によるOut-Procサーバー(EXEサーバー)の作成について

DotNETはCOMとの連携が非常に手厚くなっており、C#でCOMのIn-Procサーバ(DLL)を作るのは非常に簡単である。

ところが、標準ではOut-Procサーバー(EXE)を作成する方法は用意されていない。


しかし、COMとしての仕組みは十分に備えているため、C++(Win32)によるEXEサーバを作成する手順と同じ手順を踏むことで、C#でもOut of ProcessなEXEサーバーを実現することができる。

具体的には、以下の手順を行う。

+ EXEが提供するCOMを生成するためのクラスファクトリを、必要なクラス分だけ実装する
+ EXEは、起動したら [CoRegisterClassObject](https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms693407.aspx) でシステムにクラスファクトリを登録する
+ すべてのクラスファクトリを登録したら、[CoResumeClassObjects](https://msdn.microsoft.com/en-us/library/windows/desktop/ms692686.aspx) で、クライアントからの要求を受け付け開始する。
+ EXEサーバーは自分が不要と判断できるまで、メッセージループを回すだけの待機状態にはいる。(作成したオブジェクトがなくなるまで)
+ 自分が終了すべきと判断したら、 [CoSuspendClassObjects](https://msdn.microsoft.com/en-us/library/windows/desktop/ms691208.aspx) で受付を停止する。
  - 以後は、このEXEに対して要求が入らなくなる。(以後に新しい要求があった場合は、別のEXEが起動される。)
+ [CoRevokeClassObject](https://msdn.microsoft.com/en-us/library/windows/desktop/ms688650.aspx) でクラスファクトリの登録を解除する。
+ アプリケーションを終了する。

このあたりの流れは、

- [いちごパック COM/ActiveXの解説ページ](https://ichigopack.net/win32com/)
- [EternalWindows COM / COMアプリケーション](http://eternalwindows.jp/com/comapp/comapp00.html)

などが詳しい。



また、DotNETで作成したCOMは ```regasm``` ツールによってレジストリにCOM情報を登録するが、
標準ではDLLのIn-Procサーバーを想定したレジストリが出力される。

そこで、[ComRegisterFunction](https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.comregisterfunctionattribute.aspx) 属性を使って、COMのメソッドでレジストリ登録時の処理をオーバーライドする。

ここで、COMの種別を ```InprocServer32``` から ```LocalServer32``` に変更することで、EXEのCOMとして起動できるようになる。


また、前述した```COM Elevation Moniker```のためのレジストリエントリも、ここで追加する。


# 手順

以下、```Visual Studio Express 2017 for Windows Desktop```でC#でUACの昇格可能なEXE-COMを作成する手順を示す。

## 手順1: プロジェクトの選択

プロジェクトはコンソールとする。

実際にはメッセージループをもつWindowsフォームのEXEとなるが、フォーム画面は1つも必要なく、かわりにログメッセージ類を画面に表示させたいので、コンソールを選んでおく。

(コンソール画面が必要ないのならば、あとから出力の種類を「Windows アプリケーション」に戻しておけばよい。)

また、プラットフォームは ```x86``` (32ビット版) に固定しておく。

DLLの場合は呼び出し元のEXEと同じプラットフォームのバイナリを用意する必要があるので、32ビット版と64ビット版の2つが必要となるが、
EXEの場合は、もとからプロセス間通信でやりとりするため、呼び出し元が32ビットであろうが64ビットであろうが、どちらか1つあれば十分である。

なので、とりあえず32ビットでビルドしておけば、どこのマシンでも動くことができるであろう。


## 手順2: 参照設定でWindowsFormを指定する

コンソールアプリとしてプロジェクトを作成したが、実際にはメッセージループをまわすWindowsFormの仕組みを使う。

また、```app.config```ファイルから設定を読み込みたいので、

参照設定では

- ```System.Windows.Forms```
- ```System.Configuration```

の2つのアセンブリ参照を追加しておく。


## 手順3: 目的となるCOMオブジェクトを定義する。

COM定義用のファイルを作成する。

C#でのCOMの定義方法は、基本的にはDLLのCOMの場合と変わらない。

```cs
    /// <summary>
    /// 独自COMインターフェイスの定義
    /// </summary>
    [Guid("8CA4F6A2-4BCC-4642-B14A-C2B52E8B3DB6"), ComVisible(true)]
    public interface IMyElevationOutProcSrv
    {
        string Name { get; set; }
        void ShowHello();
    }

    [Guid("FFD359DD-D03C-4573-9986-FE5E6BDC3A29"), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _MyElevationOutProcSrvEvents
    {
        [DispId(1)]
        void NamePropertyChanging(string NewValue, ref bool Cancel);

        [DispId(2)]
        void NamePropertyChanged(string NewValue);
    }

    /// <summary>
    /// 独自のCOMオブジェクトの実装
    /// </summary>
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyElevationOutProcSrv")]
    [Guid("7AEFA37C-4494-4AE8-9378-0157A0B919AE"), ComVisible(true)]
    [ComSourceInterfaces(typeof(_MyElevationOutProcSrvEvents))] // イベント
    public class MyElevationOutProcSrv : IMyElevationOutProcSrv
    {
        private string _Name = "PiyoPiyo";

        public string Name
        {
            get
            {
                return _Name;
            }
            set
            {
                bool cancel = false;
                NamePropertyChanging?.Invoke(value, ref cancel);
                if (!cancel)
                {
                    _Name = value;
                    NamePropertyChanged?.Invoke(value);
                }
            }
        }

        public MyElevationOutProcSrv()
        {
            MyApplicationContext.Current.IncrementCount();
        }

        ~MyElevationOutProcSrv()
        {
            MyApplicationContext.Current.DecrementCount();
        }

        public void ShowHello()
        {
            Console.WriteLine("Hello, {0}!", Name);
        }

        /// <summary>
        /// NamePropertyChangingのイベント用のデリゲート(イベントソースの定義と一致していること)
        /// </summary>
        [ComVisible(false)]
        public delegate void NamePropertyChangingDelegate(string NewValue, ref bool Cancel);

        /// <summary>
        /// NamePropertyCHangedのイベント用のデリゲート(イベントソースの定義と一致していること)
        /// </summary>
        [ComVisible(false)]
        public delegate void NamePropertyChangedDelegate(string NewValue);

        /// <summary>
        /// NamePropertyChangingのイベント(イベントソースの定義と一致していること)
        /// </summary>
        public event NamePropertyChangingDelegate NamePropertyChanging;

        /// <summary>
        /// NamePropertyCHangedのイベント(イベントソースの定義と一致していること)
        /// </summary>
        public event NamePropertyChangedDelegate NamePropertyChanged;

        #region レジストリ登録
        /// <summary>
        /// Regasmツールで登録するCOMのレジストリ。
        /// 既定ではInProc(DLL)用のため、ここでLocalServer(EXE)用に修正する。
        /// </summary>
        /// <param name="typ"></param>
        [ComRegisterFunction()]
        public static void Register(Type typ)
        {
            // このEXEのフルパスを取得する
            var assembly = Assembly.GetExecutingAssembly();
            string exePath = assembly.Location;

            // アセンブリのGUIDをtypelibのIDとする。
            var attribute = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute), true)[0];
            string libid = attribute.Value;

            using (var keyCLSID = Registry.ClassesRoot.OpenSubKey(
                @"CLSID\" + typ.GUID.ToString("B"), true)) // 書き込み可能として開く
            {
                // InprocServer32を消す
                keyCLSID.DeleteSubKeyTree("InprocServer32");

                // かわりにLocalServer32とする。
                using (var subkey = keyCLSID.CreateSubKey("LocalServer32"))
                {
                    // この実行ファイル(*.exe)へのパスを登録する
                    subkey.SetValue("", exePath, RegistryValueKind.String);
                }

                // ↓ タイプライブラリの登録も行う場合
                using (var subkey = keyCLSID.CreateSubKey("TypeLib"))
                {
                    // このアセンブリのGUID(LIBID)を登録する
                    subkey.SetValue("", libid, RegistryValueKind.String);
                }

                // ↓ ここから、UACのCOM昇格可能にするための設定

                // LocalizedString 
                // "@" + EXEのフルパス + ",-" + 文字列リソース番号で文字列リソースを指定する。
                keyCLSID.SetValue("LocalizedString",
                    "@" + exePath + ",-101",
                    RegistryValueKind.String);

                // Elevation
                using (var subkey = keyCLSID.CreateSubKey("Elevation"))
                {
                    // この実行ファイル(*.exe)へのパスを登録する
                    subkey.SetValue("Enabled", 1, RegistryValueKind.DWord);
                }
            }
        }

        [ComUnregisterFunction()]
        public static void Unregister(Type t)
        {
            // レジストリエントリの削除
            Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\" + t.GUID.ToString("B"));
        }
        #endregion
    }
```

ただし、```ComRegisterFunction```, ```ComUnregisterFunction``` 属性のついたメソッドがある。

これは前述のとおり、```regasm``` ツールでCOMのレジストリ情報を登録・登録解除するときに呼び出される処理をオーバーライドするものである。

ここには、```COM Elevation Moniker``` のためのレジストリ設定も含まれている。

コンストラクタとデストラクタで

```cs
MyApplicationContext.Current.IncrementCount();
MyApplicationContext.Current.DecrementCount();
```

というカウンタをとっているところがあるが、これが、このEXEで生成したオブジェクトの残存数をカウントするためのものである。

このカウンタが0になって一定時間経過したら、このEXEサーバーは不要になったものとみなして終了するように実装する。


## 手順4: クラスファクトリの実装

DLLサーバの場合はDotNETフレームワーク側でクラスファクトリの役割が暗黙で行われていたので用意する必要はなかったのだが、

EXEサーバの場合はWin32で明示的にクラスファクトリを登録する都合上、自前のクラスファクトリを用意する必要がある。

クラスファクトリは以下の定義となる。

```cs
    /// <summary>
    /// COMのクラスファクトリ
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/ms694364(v=vs.85).aspx
    /// </summary>
    [ComImport, ComVisible(false),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
    Guid("00000001-0000-0000-C000-000000000046")]
    public interface IClassFactory
    {
        IntPtr CreateInstance([In] IntPtr pUnkOuter, [In] ref Guid riid);

        void LockServer([In] bool fLock);
    }
```
これを自分のCOMオブジェクト用のクラスファクトリとして実装する。

```cs
    /// <summary>
    /// MyElevationOutProcSrvオブジェクトのファクトリ
    /// </summary>
    public class MyElevationOutProcSrvFactory : IClassFactory
    {
        #region WIN32定義
        /// <summary>
        /// IDispatchのGUID
        /// </summary>
        public static readonly Guid GUID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        /// <summary>
        /// IUnknownのGUID
        /// </summary>
        public static readonly Guid GUID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        [Flags]
        public enum CLSCTX : uint
        {
            INPROC_SERVER = 0x1,
            INPROC_HANDLER = 0x2,
            LOCAL_SERVER = 0x4,
            INPROC_SERVER16 = 0x8,
            REMOTE_SERVER = 0x10,
            INPROC_HANDLER16 = 0x20,
            RESERVED1 = 0x40,
            RESERVED2 = 0x80,
            RESERVED3 = 0x100,
            RESERVED4 = 0x200,
            NO_CODE_DOWNLOAD = 0x400,
            RESERVED5 = 0x800,
            NO_CUSTOM_MARSHAL = 0x1000,
            ENABLE_CODE_DOWNLOAD = 0x2000,
            NO_FAILURE_LOG = 0x4000,
            DISABLE_AAA = 0x8000,
            ENABLE_AAA = 0x10000,
            FROM_DEFAULT_CONTEXT = 0x20000,
            ACTIVATE_32_BIT_SERVER = 0x40000,
            ACTIVATE_64_BIT_SERVER = 0x80000
        }

        [Flags]
        public enum REGCLS : uint
        {
            SINGLEUSE = 0,
            MULTIPLEUSE = 1,
            MULTI_SEPARATE = 2,
            SUSPENDED = 4,
            SURROGATE = 8,
        }

        [DllImport("ole32.dll", PreserveSig = false)] // HRESULTの戻り値を例外として受け取る
        protected static extern UInt32 CoRegisterClassObject(
            [In] ref Guid rclsid,
            [MarshalAs(UnmanagedType.Interface), In] IClassFactory pUnk,
            [In] CLSCTX dwClsContext,
            [In] REGCLS flags);

        [DllImport("ole32.dll", PreserveSig = false)] // HRESULTの戻り値を例外として受け取る
        static extern void CoRevokeClassObject([In] UInt32 dwRegister);
        #endregion

        #region IClassFactoryの実装
        public IntPtr CreateInstance([In] IntPtr pUnkOuter, [In] ref Guid riid)
        {
            if (pUnkOuter != IntPtr.Zero)
            {
                // アグリゲーションはサポートしていない
                Marshal.ThrowExceptionForHR(unchecked((int)0x80040110)); // CLASS_E_NOAGGREGATION
            }
            else if (riid == typeof(IMyElevationOutProcSrv).GUID ||
                riid == GUID_IDispatch || riid == GUID_IUnknown)
            {
                // IMyElevationOutProcSrv, IDispatch, IUnknownのいずれかである場合はオブジェクトを生成して返す
                var inst = new MyElevationOutProcSrv();
                return Marshal.GetComInterfaceForObject(inst, typeof(IMyElevationOutProcSrv));
            }

            // サポート外のインターフェイスが要求された場合はエラーとする
            Marshal.ThrowExceptionForHR(unchecked((int)0x80004002)); // E_NOINTERFACE
            throw new InvalidCastException(); // (E_NOINTERFACE同等)
        }

        public void LockServer([In] bool fLock)
        {
            if (fLock)
            {
                MyApplicationContext.Current.IncrementCount();
            }
            else
            {
                MyApplicationContext.Current.DecrementCount();
            }
        }
        #endregion 

        /// <summary>
        /// 登録されたクラスオブジェクトレジスターを識別するクッキー
        /// </summary>
        private uint _cookieClassObjRegister;

        /// <summary>
        /// クラスファクトリをシステムに登録する。
        /// (登録した段階ではサスペンドされている。)
        /// </summary>
        public void CoRegisterClassObject()
        {
            // COMファクトリを登録する
            var clsid = typeof(MyElevationOutProcSrv).GUID;
            _cookieClassObjRegister = CoRegisterClassObject(
                ref clsid,   // 登録するCLSID
                this,        // CLSIDをインスタンス化するクラスファクトリ
                CLSCTX.LOCAL_SERVER, // ローカルサーバーとして実行
                REGCLS.MULTIPLEUSE | REGCLS.SUSPENDED); // 複数利用可・停止状態で作成
        }

        /// <summary>
        /// クラスファクトリの登録解除を行う。
        /// </summary>
        public void CoRevokeClassObject()
        {
            if (_cookieClassObjRegister != 0)
            {
                CoRevokeClassObject(_cookieClassObjRegister);
                _cookieClassObjRegister = 0;
            }
        }
    }
```
クラスファクトリでは要求されたインターフェイスのIIDが```IUnknown```, ```IDispatch```, もしくは自分のインターフェイスのGUIDと一致すれば、
オブジェクトを生成して、それを ``` Marshal.GetComInterfaceForObject()``` を使うことで、
DotNETのオブジェクトからCOM用のインターフェイスのポインタを取得して、それを返す。


```CoRegisterClassObject``` と ```CoRevokeClassObject``` は、
アプリケーションの開始、終了時にクラスファクトリのシステムへの登録と登録解除のために呼び出されるものである。


## 手順5: アプリケーションのエントリポイントの作成

このEXEはCOMサーバーであるため、

- フォームを持たない。(ユーザーの操作で終了するわけではない。)
- 起動後は、COMクライアントからの要求を待ち受けする以外は何もしない。
  - COMのマーシャリングのためのメッセージループが必要
- アクティブな残存オブジェクトがなくなったら終了する

という、普通のフォームアプリとは、すこし違った動きをする必要がある。

このため、独自の「アプリケーションコンテキスト」を用意する。


アプリケーションは、以下のような形で開始される。

```cs
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
```

独自のアプリケーションコンテキストを生成して、設定ファイルから設定値を取り込んだら、それを ```Application.Run()``` に渡す。

アプリケーションコンテキストでは、COMのクラスファクトリ登録等々の設定後に、メッセージループに入る。

```Application.Exit``` の呼び出しをもってメッセージループは終了され、この ```Main``` メソッドに戻ってきて、アプリケーションが終了となる。



## 手順6: 独自アプリケーションコンテキストの作成

独自のアプリケーションコンテキストは、以下のような処理を行う。

+ クラスファクトリをシステムに登録する
+ クラスファクトリのクライアントからの受け入れを開始する
+ タイマーを使って、定期的に残存オブジエクト数をカウントする
  - 明示的にGCを呼び出してデストラクタを動かして残存オブジェクト数を確定させる
  - 残存オブジェクト数が0になって一定時間経過したら終了処理を行う
+ 終了する場合は、まず、クライアントからの受け入れをすべて停止して、システムからクラスファクトリの登録を解除する
  - (登録解除後、あらためて残存数が0であれば)メッセージループを終了(```Application.Exit```)する。

以上のものを実装すると、以下のようになる。

```cs
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
```

いくつか、些細な注意点がある。

- タイマーで定期的にGCを行う必要がある。
  - ```Application.Idle```イベントで、アイドル時にGCすることも試してみたが、適切ではなかった。
    - オブジェクトが全てReleaseされるとマーシャリングのためのメッセージも発生しなくなり、Idleイベントも発生しなくなる。
- **過去に1つ以上のオブジェクトを生成済み** の場合のみ、残存オブジェクトが0になったあとのタイムアウト判定を行う。
  - デバック実行等で、COMによりEXEが起動したのに、オブジェクトが1つも作成されないうちにタイムアウトしてアプリケーションを終了すると、システムが混乱する為。
  - COMによってEXEが起動した場合は、引数として [**-Embedding** フラグ](https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms683844.aspx)が付与されるので、必要ならば、これで判定する。

## 手順7: Win32リソースのコンパイルとEXEへの埋め込み

手順3で、```COM Elevation Moniker``` が有効となるためには、UACの昇格ダイアログに表示するための、

```LocalizedString```キーで、**"@" + EXEのフルパス + ",-" + 文字列リソース番号** の文字列リソースを指定する必要がある。

これは、たとえば、文字列リソース番号が ```101``` の場合、リソースファイル```resource.rc``` は以下のように定義する必要がある。

```
#define IDS_STRING101   101

STRINGTABLE
BEGIN
IDS_STRING101           "MyElevationOutProcSrv"
END
```

このリソースソースは、リソースコンパイラ(rc.exe) によってコンパイルされ、```resource.res``` というファイルになる。


問題は、VC++プロジェクトの場合であれば、*.rc ファイルはリソースファイルとして認識され、コンパイルされるが、
C#プロジェクトの場合は、ソリューションエクスプローラから *.rcファイルを追加してもリソースコンパイラのソースとしては認識されない点である。


開発者コマンドプロンプトを開いて、手作業で ```rc.exe``` でコンパイルしても良いのだが、
```MSBuild``` を手動で書き換えて、C#のプロジェクト中にVC++用のリソースコンパイラのタスクを流用させてしまう手法がある。

- https://www.e-learn.cn/content/wangluowenzhang/466086
- https://msdn.microsoft.com/en-us/library/ee862475.aspx


### MyElevationOutProcSrv.csprojファイルにrcタスクとリソースファイル設定を追加する方法

*.csprojファイルを直接開いて、以下のようにタスクを追記する。

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" Condition="Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')" />
  <UsingTask TaskName="RC" AssemblyName="Microsoft.Build.CppTasks.Common, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" />
  <Target Name="ResourceCompile" BeforeTargets="BeforeCompile" Condition="'@(ResourceCompile)' != ''">
    <RC Source="@(ResourceCompile)" />
  </Target>
  <PropertyGroup>
  ....
    <Win32Resource>@(ResourceCompile->'%(RelativeDir)%(filename).res')</Win32Resource>
  </PropertyGroup>
  <ItemGroup>
    <ResourceCompile Include="resource.rc" />
  </ItemGroup>
  ....
</Project>
```

```<UsingTask>``` で、```Microsoft.Build.CppTasks.Common``` のVC++用のタスクを使用して、

```<Target Name="ResourceCompile"...><RC Source="@(ResourceCompiler)"/>...``` で、リソースコンパイラ定義部分を、RCタスクで処理する。

これにより、

```xml
    <ResourceCompile Include="resource.rc" />
```
と指定されている ```resource.rc``` ファイルは、rc.exeによってコンパイルされて、 ```resource.res``` に変換される。

この変換結果を、プロパティグループ内の

```xml
    <Win32Resource>@(ResourceCompile->'%(RelativeDir)%(filename).res')</Win32Resource>
```

に設定することで、生成されたresファイルをwin32リソースとしてEXEに取り込む設定となる。


もし、手動で ```rc.exe``` でコンパイルするのであれば、コンパイル結果の ```*.res``` ファイル名を、直接、ここに書けば良い。

```xml
    <Win32Resource>resource.res</Win32Resource>
```

なお、この箇所は画面からも確認できるが入力はできない。(ファイル名としてチェックされて、エラーになって保存できないため。)



## 手順8: regasmによるレジストリへの登録

EXEがビルドできたら、これをレジストリに登録する。

開発者コマンドプロンプトを **管理者権限** で開いて、```regasm```によりレジストリに登録する。

```
regasm /codebase /tlb Debug\MyElevationOutProcSrv.exe
```

```/codebase``` を指定する場合は厳密名をつけて署名しろ、みたいな警告がでるが、とりあえずレジストリには登録できているはずである。

```/tlb``` はタイプライブラリのレジストリへの登録を行うものである。
(COMからのイベントをハンドルしないのであれば、これは指定しなくても、メソッドやプロパティの呼び出しには支障ない。)


なお、登録解除する場合は、

```
regasm /u /codebase /tlb Debug\MyElevationOutProcSrv.exe
```

のように行う。


# 動作確認


以上で、COM昇格可能なOut-ProcなCOMサーバーが使えるようになっている。

## VBAからの動作確認

ExcelのVBAから動作確認してみる。

VBAからCOMのイベントをハンドルするためには、参照設定でタイプライブラリを指定しなければならない。


```vba
Private WithEvents obj As MyElevationOutProcSrv.MyElevationOutProcSrv

Public Sub TestCOM()
    Set obj = GetObject("Elevation:Administrator!new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
    'Set obj = GetObject("new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
    obj.name = "FooBar"
    Call obj.ShowHello
    Set obj = Nothing
End Sub

Private Sub obj_NamePropertyChanging(ByVal name As String, ByRef cancel As Boolean)
    If (MsgBox("Changing? " & name, vbYesNo, "Confirm") <> vbYes) Then
        cancel = True
    End If
End Sub

Private Sub obj_NamePropertyChanged(ByVal name As String)
    MsgBox "changed: " & name
End Sub
```

(なお、イベントをハンドルする必要がなければタイプライブラリの参照設定は不要である。)

## VBSからの動作確認

VBSからは、以下のように使える。

```vbs
Option Explicit
Dim obj

'Set obj = CreateObject("MyElevationOutProcSrv")
'Set obj = GetObject("new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
Set obj = GetObject("Elevation:Administrator!new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
WScript.ConnectObject obj, "obj_"

obj.Name = "FooBar"
obj.ShowHello()

Sub obj_NamePropertyChanging(ByVal name, ByRef cancel)
	WScript.Echo("Changing: " & name)
	' VBSからはbyrefのcancel値は返却できない。
End Sub

Sub obj_NamePropertyChanged(ByVal name)
	WScript.Echo("Changed: " & name)
End Sub
```

ただし、```COM Elevation Moniker``` を経由する場合は、```WScript.GetObject()``` ではなく、直接、```GetObject()``` を使う必要があるようである。(理由は不明)


以上、メモ終了。






