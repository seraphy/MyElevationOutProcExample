using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MyElevationOutProcSrv
{
    /// <summary>
    /// 独自COMインターフェイスの定義
    /// </summary>
    [Guid("8CA4F6A2-4BCC-4642-B14A-C2B52E8B3DB6"), ComVisible(true)]
    public interface IMyElevationOutProcSrv
    {
        string Name { get; set; }
        void ShowHello();
    }

    /// <summary>
    /// 独自COMイベントの定義
    /// </summary>
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
            // 残存オブジェクト数 +1
            MyApplicationContext.Current.IncrementCount();
        }

        ~MyElevationOutProcSrv()
        {
            // 残存オブジェクト数 -1
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
}
