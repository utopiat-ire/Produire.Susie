// Produire Susie ホストプラグイン //
// Copyright(C) 2022 utopiat.net https://github.com/utopiat-ire/
// Susie API仕様: http://www2f.biglobe.ne.jp/~kana/spi_api/index.html

using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Produire.Susie
{
	[種類(DocUrl = "/plugins/susie/susie.htm")]
	public class Susieホスト : IProduireStaticClass, IDisposable
	{
		List<SusiePlugin> pluginList = new List<SusiePlugin>();
		public Susieホスト()
		{
		}
		[自分へ]
		public bool 登録([を] string パス)
		{
			var plugin = SusiePlugin.Create(パス);
			if (plugin == null) return false;
			pluginList.Add(plugin);
			return true;
		}

		[自分("で")]
		public Bitmap 開く([を] string 画像ファイル)
		{
			var plugin = GetPluginFor(画像ファイル) as SusieImagePlugin;
			if (plugin == null) return null;
			Bitmap bmp = plugin.GetPicture(画像ファイル);
			return bmp;
		}

		[自分("で")]
		public void 抽出([から] string 書庫ファイル, [を] string 対象ファイル名, [へ] string 抽出先ファイル)
		{
			var plugin = GetPluginFor(書庫ファイル) as SusieArchivePlugin;
			if (plugin == null) return;
			var info = plugin.GetFileInfo(書庫ファイル, 対象ファイル名);
			if (info.method.Length > 0)
			{
				plugin.GetFile(書庫ファイル, info.position, 抽出先ファイル);
			}
		}

		[自分("で")]
		public byte[] 抽出([から] string 書庫ファイル, [を] string 対象ファイル名)
		{
			var plugin = GetPluginFor(書庫ファイル) as SusieArchivePlugin;
			if (plugin == null) return null;
			var info = plugin.GetFileInfo(書庫ファイル, 対象ファイル名);
			byte[] result = null;
			if (info.method.Length > 0)
			{
				result = plugin.GetFileBytes(書庫ファイル, info.position);
			}
			return result;
		}

		[自分("で"), 名詞手順("一覧")]
		public Susie書庫項目[] 一覧([の] string 書庫ファイル)
		{
			var plugin = GetPluginFor(書庫ファイル) as SusieArchivePlugin;
			if (plugin == null) return null;
			var files = plugin.GetArchiveInfo(書庫ファイル);
			return files;
		}

		internal SusiePlugin GetPluginFor(string filename)
		{
			SusiePlugin selected = null;
			using (var file = File.OpenRead(filename))
			{
				foreach (var plugin in pluginList)
				{
					if (plugin.IsSupported(filename, file.SafeFileHandle))
					{
						selected = plugin;
						break;
					}
				}
			}
			return selected;
		}
		public List<SusiePlugin> プラグイン一覧
		{
			get { return pluginList; }
		}

		public void Dispose()
		{
			foreach (var plugin in pluginList)
			{
				plugin.Dispose();
			}
		}
	}

	[種類("Susieプラグイン")]
	public class SusiePlugin : IDisposable, IProduireClass
	{
		#region API関数

		/// <summary>DLLをロード</summary>
		[DllImport("kernel32", EntryPoint = "LoadLibrary", SetLastError = true, CharSet = CharSet.Auto, ExactSpelling = false)]
		protected extern static IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPTStr)] string lpFileName);

		/// <summary>解放</summary>
		[DllImport("kernel32", EntryPoint = "FreeLibrary", SetLastError = true, ExactSpelling = true)]
		protected extern static bool FreeLibrary(IntPtr hModule);

		/// <summary>関数のアドレスを取得</summary>
		[DllImport("kernel32", EntryPoint = "GetProcAddress", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true)]
		protected extern static IntPtr GetProcAddress(IntPtr hModule, [MarshalAs(UnmanagedType.LPStr)] string lpProcName);


		[DllImport("kernel32")]
		protected extern static IntPtr LocalLock(IntPtr hMem);

		[DllImport("kernel32")]
		protected extern static bool LocalUnlock(IntPtr hMem);

		[DllImport("kernel32")]
		protected extern static IntPtr LocalFree(IntPtr hMem);

		[DllImport("kernel32")]
		[return: MarshalAs(UnmanagedType.SysUInt)]
		protected extern static UIntPtr LocalSize(IntPtr hMem);



		/// <summary>GetPluginInfo</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetPluginInfoDelegate(ushort infono, [MarshalAs(UnmanagedType.LPStr)] StringBuilder buf, ushort buflen);

		/// <summary>IsSupported</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate bool SpiIsSupportedDelegate([MarshalAs(UnmanagedType.LPStr)] string buf, SafeFileHandle dw);

		/// <summary>ConfigurationDlg</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiConfigurationDlgDelegate(IntPtr parent, ushort func);

		#endregion

		internal static SusiePlugin Create(string spiPath)
		{
			IntPtr hModule = Load(spiPath);
			string apiver = GetPluginInfo(hModule, 0);
			if (apiver.EndsWith("IN"))
				return new SusieImagePlugin(hModule);
			else if (apiver.EndsWith("AM"))
				return new SusieArchivePlugin(hModule);
			else
				return null;
		}


		protected IntPtr hModule;

		protected SusiePlugin(IntPtr hModule)
		{
			this.hModule = hModule;
		}


		#region ラッパー

		protected static IntPtr Load(string archiverDllName)
		{
			IntPtr hModule = LoadLibrary(archiverDllName);
			if (hModule == IntPtr.Zero)
			{
				string message = archiverDllName + "が利用できません。次の点を確認してください。\n";
				if (!File.Exists(archiverDllName)) message += "・" + archiverDllName + "が置かれているかどうか\n";
				if (IsX86)
					message += "・32ビット版SPIかどうか\n";
				else
					message += "・64ビット版SPIかどうか\n";
				throw new ProduireException(message + "エラーコード:" + Marshal.GetLastWin32Error().ToString());
			}
			return hModule;
		}

		protected static int GetPluginInfo(IntPtr hModule, ushort infono, StringBuilder buf, ushort buflen)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetPluginInfo");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetPluginInfoDelegate)) as SpiGetPluginInfoDelegate;
			int ret = execute(infono, buf, buflen);
			return ret;
		}

		internal bool IsSupported(string buf, SafeFileHandle dw)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "IsSupported");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiIsSupportedDelegate)) as SpiIsSupportedDelegate;
			bool ret = execute(buf, dw);
			return ret;
		}

		private int ConfigurationDlg(IntPtr parent, ushort func)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "ConfigurationDlg");
			if (funcAddr == IntPtr.Zero)
			{
				return -1;
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiConfigurationDlgDelegate)) as SpiConfigurationDlgDelegate;
			int ret = execute(parent, func);
			return ret;
		}

		internal static string GetPluginInfo(IntPtr hModule, int no)
		{
			StringBuilder buf = new StringBuilder(256);
			GetPluginInfo(hModule, (ushort)no, buf, (ushort)buf.Capacity);
			return buf.ToString();
		}

		#endregion

		[自分("が"), 手順名("対応する")]
		public bool IsSupported([を] string filename)
		{
			bool supported;
			using (var file = File.OpenRead(filename))
			{
				supported = IsSupported(filename, file.SafeFileHandle);
			}
			return supported;
		}

		[自分("にて"), 手順名("バージョン情報を", "表示する")]
		public void ShowAbout([へ, 省略] Form form)
		{
			ConfigurationDlg((form == null) ? IntPtr.Zero : form.Handle, 0);
		}

		[自分("にて"), 手順名("設定画面を", "表示する")]
		public void ShowConfig([へ, 省略] Form form)
		{
			ConfigurationDlg((form == null) ? IntPtr.Zero : form.Handle, 1);
		}

		#region プロパティ

		[設定項目("APIバージョン")]
		public string APIVersion
		{
			get
			{
				return GetPluginInfo(hModule, 0);
			}
		}

		[設定項目("名前")]
		public string PluginName
		{
			get
			{
				return GetPluginInfo(hModule, 1);
			}
		}

		[設定項目("対応フィルタ")]
		public string FilterString
		{
			get
			{
				StringBuilder filter = new StringBuilder();
				StringBuilder buf = new StringBuilder(256);
				int n = 0;
				for (; ; )
				{
					GetPluginInfo(hModule, (ushort)(2 * n + 3), buf, (ushort)buf.Capacity);
					if (buf.Length == 0) break;
					if (filter.Length > 0) filter.Append("|");
					filter.Append(buf.ToString());
					GetPluginInfo(hModule, (ushort)(2 * n + 2), buf, (ushort)buf.Capacity);
					filter.Append("|");
					filter.Append(buf.ToString());
					n++;
				}
				return filter.ToString();
			}
		}

		[設定項目("対応ファイルパターン")]
		public string AllPattanString
		{
			get
			{
				StringBuilder filter = new StringBuilder();
				StringBuilder buf = new StringBuilder(256);
				int n = 0;
				for (; ; )
				{
					if (filter.Length > 0) filter.Append(";");
					GetPluginInfo(hModule, (ushort)(2 * n + 2), buf, (ushort)buf.Capacity);
					if (buf.Length == 0) break;
					filter.Append(buf.ToString());
					n++;
				}
				return filter.ToString();
			}
		}

		#endregion

		public void Dispose()
		{
			if (hModule != IntPtr.Zero) FreeLibrary(hModule);
		}

		public static bool IsX86
		{
			get { return (IntPtr.Size * 8) == 32; }
		}

	}

	[種類("Susie画像プラグイン")]
	public class SusieImagePlugin : SusiePlugin
	{
		internal protected SusieImagePlugin(IntPtr hModule)
			: base(hModule)
		{

		}

		/// <summary>GetPictureInfo</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetPictureInfoDelegate([MarshalAs(UnmanagedType.LPStr)] string buf, int len, ushort flag, out PictureInfo lpInfo);

		/// <summary>GetPicture</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetPictureDelegate([MarshalAs(UnmanagedType.LPStr)] string buf, int len, ushort flag, out IntPtr pHBInfo, out IntPtr pHBm, IntPtr lpPrgressCallback, int lData);

		/// <summary>GetPreview</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetPreviewDelegate([MarshalAs(UnmanagedType.LPStr)] string buf, int len, ushort flag, out IntPtr pHBInfo, out IntPtr pHBm, IntPtr lpPrgressCallback, int lData);


		[StructLayout(LayoutKind.Sequential, Pack = 1)]
		public struct PictureInfo
		{
			/// <summary>
			/// 画像を展開する位置
			/// </summary>
			public int left;
			/// <summary>
			/// 画像を展開する位置
			/// </summary>
			public int top;
			/// <summary>
			/// 画像の幅
			/// </summary>
			public int width;
			/// <summary>
			/// 画像の高さ
			/// </summary>
			public int height;
			/// <summary>
			/// 画素の水平方向密度
			/// </summary>
			public byte xDensity;
			/// <summary>
			/// 画素の垂直方向密度
			/// </summary>
			public byte yDensity;
			/// <summary>
			/// 画素当たりのbit数
			/// </summary>
			public byte colorDepth;
			/// <summary>
			/// 画像内のテキスト情報
			/// </summary>
			public IntPtr hInfo;
		}

		private int GetPictureInfo(string buf, int len, ushort flag, out PictureInfo lpInfo)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetPictureInfo");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetPictureInfoDelegate)) as SpiGetPictureInfoDelegate;
			int ret = execute(buf, len, flag, out lpInfo);
			return ret;
		}

		private int GetPicture(string buf, int len, ushort flag, out IntPtr pHBInfo, out IntPtr pHBm, IntPtr lpPrgressCallback, int lData)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetPicture");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetPictureDelegate)) as SpiGetPictureDelegate;
			int ret = execute(buf, len, flag, out pHBInfo, out pHBm, lpPrgressCallback, lData);
			return ret;
		}

		private int GetPreview(string buf, int len, ushort flag, out IntPtr pHBInfo, out IntPtr pHBm, IntPtr lpPrgressCallback, int lData)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetPreview");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetPreviewDelegate)) as SpiGetPreviewDelegate;
			int ret = execute(buf, len, flag, out pHBInfo, out pHBm, lpPrgressCallback, lData);
			return ret;
		}

		#region メソッド

		public Bitmap GetPicture(string filename)
		{
			int ret = GetPicture(filename, 0, 0, out IntPtr infoHandle, out IntPtr dataHandle, IntPtr.Zero, 0);
			if (ret != 0) return null;
			Bitmap bmp = CreateBitmap(infoHandle, dataHandle);
			return bmp;
		}

		private static Bitmap CreateBitmap(IntPtr infoHandle, IntPtr dataHandle)
		{
			IntPtr struPtr1 = LocalLock(infoHandle);
			IntPtr struPtr2 = LocalLock(dataHandle);
			Bitmap bmp;
			try
			{
				var info = (BITMAPINFOHEADER)Marshal.PtrToStructure(struPtr1, typeof(BITMAPINFOHEADER));
				PixelFormat format;
				if (info.biBitCount == 8)
					format = PixelFormat.Format8bppIndexed;
				else if (info.biBitCount == 16)
					format = PixelFormat.Format16bppRgb555;
				else if (info.biBitCount == 32)
					format = PixelFormat.Format32bppArgb;
				else
					format = PixelFormat.Format24bppRgb;
				bmp = new Bitmap(info.biWidth, info.biHeight, (info.biWidth * info.biBitCount + 31) / 32 * 4, format, struPtr2);
				bmp.RotateFlip(RotateFlipType.RotateNoneFlipY);
			}
			finally
			{
				LocalUnlock(infoHandle);
				LocalUnlock(dataHandle);
				Marshal.FreeHGlobal(struPtr1);
				Marshal.FreeHGlobal(struPtr2);
			}
			return bmp;
		}

		internal Bitmap GetPreview(string filename)
		{
			int ret = GetPreview(filename, 0, 0, out IntPtr infoHandle, out IntPtr dataHandle, IntPtr.Zero, 0);
			if (ret == -1)
			{
				ret = GetPicture(filename, 0, 0, out infoHandle, out dataHandle, IntPtr.Zero, 0);
			}
			if (ret != 0) return null;
			Bitmap bmp = CreateBitmap(infoHandle, dataHandle);
			return bmp;
		}

		internal PictureInfo GetPictureInfo(string filename)
		{
			GetPictureInfo(filename, 0, 0, out PictureInfo info);
			if (info.hInfo != IntPtr.Zero) Marshal.FreeHGlobal(info.hInfo);
			return info;
		}
		#endregion
	}

	[種類("Susie書庫プラグイン")]
	public class SusieArchivePlugin : SusiePlugin
	{
		internal protected SusieArchivePlugin(IntPtr hModule)
			: base(hModule)
		{

		}

		/// <summary>GetArchiveInfo</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetArchiveInfoDelegate([MarshalAs(UnmanagedType.LPStr)] string buf, uint len, ushort flag, out IntPtr lphInf);

		/// <summary>GetFileInfo</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetFileInfoDelegate([MarshalAs(UnmanagedType.LPStr)] string buf, uint len, string filename, ushort flag, out SpiFileInfo lpInfo);

		/// <summary>GetFile</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetFileDelegate([MarshalAs(UnmanagedType.LPStr)] string src, uint len, string dest, ushort flag, IntPtr lpPrgressCallback, int lData);

		/// <summary>GetFile</summary>
		[UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
		delegate int SpiGetFileBytesDelegate([MarshalAs(UnmanagedType.LPStr)] string src, uint len, out IntPtr dest, ushort flag, IntPtr lpPrgressCallback, int lData);

		public Susie書庫項目[] GetArchiveInfo(string archiveFile)
		{
			List<Susie書庫項目> list = new List<Susie書庫項目>();
			GetArchiveInfo(archiveFile, 0, 0, out IntPtr intPtr);
			for (; ; )
			{
				var info = (SpiFileInfo)Marshal.PtrToStructure(intPtr, typeof(SpiFileInfo));
				intPtr += Marshal.SizeOf(typeof(SpiFileInfo));
				if (string.IsNullOrEmpty(info.method)) break;
				list.Add(new Susie書庫項目(info));
			}
			return list.ToArray();
		}

		private int GetArchiveInfo(string archiveFile, uint v1, ushort v2, out IntPtr buflen)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetArchiveInfo");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetArchiveInfoDelegate)) as SpiGetArchiveInfoDelegate;
			int ret = execute(archiveFile, v1, v2, out buflen);
			return ret;
		}

		public SpiFileInfo GetFileInfo(string archiveName, string filename)
		{
			GetFileInfo(archiveName, 0, filename, 0, out SpiFileInfo info);
			return info;
		}

		private int GetFileInfo(string archiveName, uint v1, string filename, ushort v2, out SpiFileInfo info)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetFileInfo");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetFileInfoDelegate)) as SpiGetFileInfoDelegate;
			int ret = execute(archiveName, v1, filename, v2, out info);
			return ret;
		}

		public int GetFile(string filename, uint position, string extractFolder)
		{
			int ret = GetFile(filename, position, extractFolder, IntPtr.Zero, 0);
			return ret;
		}

		private int GetFile(string filename, uint position, string extractFolder, IntPtr lpPrgressCallback, int lData)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetFile");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetFileDelegate)) as SpiGetFileDelegate;
			int ret = execute(filename, position, extractFolder, 0, lpPrgressCallback, lData);
			return ret;
		}
		public byte[] GetFileBytes(string filename, uint position)
		{
			byte[] data = GetFileBytes(filename, position, IntPtr.Zero, 0);
			return data;
		}
		private byte[] GetFileBytes(string filename, uint position, IntPtr lpPrgressCallback, int lData)
		{
			IntPtr funcAddr = GetProcAddress(hModule, "GetFile");
			if (funcAddr == IntPtr.Zero)
			{
				throw new InvalidOperationException("関数のアドレスを取得できませんでした。");
			}
			var execute = Marshal.GetDelegateForFunctionPointer(funcAddr, typeof(SpiGetFileBytesDelegate)) as SpiGetFileBytesDelegate;
			IntPtr dest = IntPtr.Zero;
			byte[] data;
			try
			{
				int ret = execute(filename, position, out dest, 256, lpPrgressCallback, lData);
				if (dest == IntPtr.Zero) return null;
				IntPtr source = LocalLock(dest);
				int size = (int)LocalSize(dest);
				data = new byte[size];
				Marshal.Copy(source, data, 0, size);
			}
			finally
			{
				if (dest != IntPtr.Zero)
				{
					LocalUnlock(dest);
					Marshal.FreeHGlobal(dest);
				}
			}
			return data;
		}
	}

	[StructLayout(LayoutKind.Sequential)]
	internal struct BITMAPINFOHEADER
	{
		public uint biSize;
		public int biWidth;
		public int biHeight;
		public ushort biPlanes;
		public ushort biBitCount;
		public uint biCompression;
		public uint biSizeImage;
		public int biXPelsPerMeter;
		public int biYPelsPerMeter;
		public uint biClrUsed;
		public uint biClrImportant;
	}

	[StructLayout(LayoutKind.Sequential, Pack = 1)]
	public struct SpiFileInfo
	{
		/// <summary>
		/// 圧縮法の種類
		/// </summary>
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
		public string method; //8       
		/// <summary>
		/// ファイル上での位置
		/// </summary>
		public uint position;
		/// <summary>
		/// 圧縮されたサイズ
		/// </summary>
		public uint compsize;
		/// <summary>
		/// 元のファイルサイズ
		/// </summary>
		public uint filesize;
		/// <summary>
		/// ファイルの更新日時
		/// </summary>
		public uint timestamp;
		/// <summary>
		/// 相対パス
		/// </summary>
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 200)]
		public string path;  //200                 
		/// <summary>
		/// ファイルネーム
		/// </summary>
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 200)]
		public string filename; //200              
		/// <summary>
		/// CRC
		/// </summary>
		public uint crc;
	}
	public struct Susie書庫項目 : IProduireClass
	{
		SpiFileInfo info;
		public Susie書庫項目(SpiFileInfo info)
		{
			this.info = info;
		}

		#region 設定項目

		public string 圧縮方法 { get { return info.method; } }
		public uint 位置 { get { return info.position; } }
		public uint 圧縮後サイズ { get { return info.compsize; } }
		public uint サイズ { get { return info.filesize; } }
		public uint 更新日時 { get { return info.timestamp; } }
		public string パス { get { return info.path; } }
		public string ファイル名 { get { return info.filename; } }
		public uint CRC { get { return info.crc; } }

		#endregion
	}

}
