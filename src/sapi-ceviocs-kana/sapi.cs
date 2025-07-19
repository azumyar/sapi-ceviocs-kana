using Microsoft.VisualBasic;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using TTSEngineLib;

namespace Yarukizero.Net.Sapi.CeVioKana;

internal static class GuidConst {
	public const string InterfaceGuid = "7B01E0C7-ABF1-4171-8226-3973F72BE777";
	public const string ClassGuid = "6CC9C867-7FF6-4637-8B7D-44ABE509CA46";
}

[ComVisible(true)]
[Guid(GuidConst.InterfaceGuid)]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ICeVioKanaTTSEngine : ISpTTSEngine, ISpObjectWithToken { }

[ComVisible(true)]
[Guid(GuidConst.ClassGuid)]
public class CeVioKanaTTSEngine : ICeVioKanaTTSEngine {
	private const ushort WAVE_FORMAT_PCM = 1;

	private static readonly Guid SPDFID_WaveFormatEx = new Guid("C31ADBAE-527F-4ff5-A230-F62BB61FF70C");
	private static readonly Guid SPDFID_Text = new Guid("7CEEF9F9-3D13-11d2-9EE7-00C04F797396");

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	private static extern IntPtr CreateFile(string pszFileName, int dwAccess, int dwShare, IntPtr psa, int dwCreatDisposition, int dwFlagsAndAttributes, IntPtr hTemplate);

	[DllImport("kernel32.dll")]
	private static extern bool ReadFile(IntPtr hFile, byte[] pBuffer, int nNumberOfBytesToRead, out int pNumberOfBytesRead, IntPtr pOverlapped);

	[DllImport("kernel32.dll")]
	private static extern int GetFileSize(IntPtr hFile, IntPtr pFileSizeHigh);
	[DllImport("kernel32.dll")]
	private static extern bool CloseHandle(IntPtr hObject);

	[DllImport("kernel32.dll")]
	private static extern int WaitForSingleObject(IntPtr hHandle, int dwMilliseconds);
	private const int WAIT_TIMEOUT = 0x102;
	private const int GENERIC_WRITE = 0x40000000;
	private const int GENERIC_READ = unchecked((int)0x80000000);
	private const int FILE_SHARE_READ = 0x00000001;
	private const int FILE_SHARE_WRITE = 0x00000002;
	private const int FILE_SHARE_DELETE = 0x00000004;
	private const int CREATE_NEW = 1;
	private const int CREATE_ALWAYS = 2;
	private const int OPEN_EXISTING = 3;
	private const int OPEN_ALWAYS = 4;
	private const int TRUNCATE_EXISTING = 5;

	[Flags]
	enum SPVESACTIONS {
		SPVES_CONTINUE = 0,
		SPVES_ABORT = (1 << 0),
		SPVES_SKIP = (1 << 1),
		SPVES_RATE = (1 << 2),
		SPVES_VOLUME = (1 << 3)
	}

	enum SPEVENTENUM {
		SPEI_UNDEFINED = 0,
		SPEI_START_INPUT_STREAM = 1,
		SPEI_END_INPUT_STREAM = 2,
		SPEI_VOICE_CHANGE = 3,
		SPEI_TTS_BOOKMARK = 4,
		SPEI_WORD_BOUNDARY = 5,
		SPEI_PHONEME = 6,
		SPEI_SENTENCE_BOUNDARY = 7,
		SPEI_VISEME = 8,
		SPEI_TTS_AUDIO_LEVEL = 9,
		SPEI_TTS_PRIVATE = 15,
		SPEI_MIN_TTS = 1,
		SPEI_MAX_TTS = 15,
		SPEI_END_SR_STREAM = 34,
		SPEI_SOUND_START = 35,
		SPEI_SOUND_END = 36,
		SPEI_PHRASE_START = 37,
		SPEI_RECOGNITION = 38,
		SPEI_HYPOTHESIS = 39,
		SPEI_SR_BOOKMARK = 40,
		SPEI_PROPERTY_NUM_CHANGE = 41,
		SPEI_PROPERTY_STRING_CHANGE = 42,
		SPEI_FALSE_RECOGNITION = 43,
		SPEI_INTERFERENCE = 44,
		SPEI_REQUEST_UI = 45,
		SPEI_RECO_STATE_CHANGE = 46,
		SPEI_ADAPTATION = 47,
		SPEI_START_SR_STREAM = 48,
		SPEI_RECO_OTHER_CONTEXT = 49,
		SPEI_SR_AUDIO_LEVEL = 50,
		SPEI_SR_RETAINEDAUDIO = 51,
		SPEI_SR_PRIVATE = 52,
		SPEI_ACTIVE_CATEGORY_CHANGED = 53,
		SPEI_RESERVED5 = 54,
		SPEI_RESERVED6 = 55,
		SPEI_MIN_SR = 34,
		SPEI_MAX_SR = 55,
		SPEI_RESERVED1 = 30,
		SPEI_RESERVED2 = 33,
		SPEI_RESERVED3 = 63
	}

	private const ulong SPFEI_FLAGCHECK = (1u << (int)SPEVENTENUM.SPEI_RESERVED1) | (1u << (int)SPEVENTENUM.SPEI_RESERVED2);
	private const ulong SPFEI_ALL_TTS_EVENTS = 0x000000000000FFFEul | SPFEI_FLAGCHECK;
	private const ulong SPFEI_ALL_SR_EVENTS = 0x003FFFFC00000000ul | SPFEI_FLAGCHECK;
	private const ulong SPFEI_ALL_EVENTS = 0xEFFFFFFFFFFFFFFFul;

	private ulong SPFEI(SPEVENTENUM SPEI_ord) => (1ul << (int)SPEI_ord) | SPFEI_FLAGCHECK;

	enum SPEVENTLPARAMTYPE {
		SPET_LPARAM_IS_UNDEFINED = 0,
		SPET_LPARAM_IS_TOKEN = (SPET_LPARAM_IS_UNDEFINED + 1),
		SPET_LPARAM_IS_OBJECT = (SPET_LPARAM_IS_TOKEN + 1),
		SPET_LPARAM_IS_POINTER = (SPET_LPARAM_IS_OBJECT + 1),
		SPET_LPARAM_IS_STRING = (SPET_LPARAM_IS_POINTER + 1)
	}

	private static readonly string KeyCeVioCast = "x-cevio-cast";
	private static readonly string KeyCeVioVolume = "x-cevio-volume";
	private static readonly string KeyCeVioSpeed = "x-cevio-speed";
	private static readonly string KeyCeVioTone = "x-cevio-tone";
	private static readonly string KeyCeVioToneScale = "x-cevio-tone-scale";
	private static readonly string KeyCeVioAlpha = "x-cevio-alpha";
	private static readonly string KeyCeVioComponents = "x-cevio-components";
	private static readonly string KeyConvertKana = "x-kana";

	private static readonly uint DefaultVolume = 50;
	private static readonly uint DefaultSpeed = 50;
	private static readonly uint DefaultTone = 50;
	private static readonly uint DefaultToneScale = 50;
	private static readonly uint DefaultAlpha = 50;
	private static readonly string DefaultComponents = "";
	private static readonly int DefaultConvertKana = 1;

	private ISpObjectToken? token;
	dynamic? cevio = default;
	dynamic? talker = default;
	private string cevioCast = "";
	private uint cevioVolume = DefaultVolume;
	private uint cevioSpeed = DefaultSpeed;
	private uint cevioTone = DefaultTone;
	private uint cevioToneScale = DefaultToneScale;
	private uint cevioAlpha = DefaultAlpha;
	private IEnumerable<(string Name, uint Value)> cevioComponents = Array.Empty<(string, uint)>();
	private string? convertKana = null;
	private System.Media.SoundPlayer? player = null;

	public void Speak(uint dwSpeakFlags, ref Guid rguidFormatId, ref WAVEFORMATEX pWaveFormatEx, ref SPVTEXTFRAG pTextFragList, ISpTTSEngineSite pOutputSite) {
		System.Diagnostics.Debug.Assert(this.cevio is not  null);
		System.Diagnostics.Debug.Assert(this.talker is not  null);
		static bool deleteTmp(string tmpWaveFile) {
			try {
				if(File.Exists(tmpWaveFile)) {
					File.Delete(tmpWaveFile);
				}
				return true;
			}
			catch(IOException) {
				return false;
			}
		}
		static uint output(ISpTTSEngineSite output, byte[] data) {
			var pWavData = IntPtr.Zero;
			try {
				if(data.Length == 0) {
					output.Write(pWavData, 0u, out var written);
					return written;
				} else {
					pWavData = Marshal.AllocCoTaskMem(data.Length);
					Marshal.Copy(data, 0, pWavData, data.Length);
					output.Write(pWavData, (uint)data.Length, out var written);
					return written;
				}
			}
			finally {
				if(pWavData != IntPtr.Zero) {
					Marshal.FreeCoTaskMem(pWavData);
				}
			}
		}
		void play(string resourceName) {
			if(this.player == null) {
				this.player = new System.Media.SoundPlayer();
			}
			player.Stream = typeof(CeVioKanaTTSEngine)
				.Assembly
				.GetManifestResourceStream(resourceName);
			player.Play();
		}

		if(rguidFormatId == SPDFID_Text) {
			return;
		}

		var optSpeed = 1d;
		{
			pOutputSite.GetRate(out var spd);
			optSpeed = Math.Max(Math.Min(1d, spd / 10d), -1d) + 1;
		}
		var optVolume = 1f;
		{
			pOutputSite.GetVolume(out var vol);
			optVolume = vol / 100f;
		}

		var volume = (uint)((this.cevioVolume / 100.0) * optVolume * 100);
		var speed = (uint)((this.cevioSpeed / 100.0) * optSpeed * 100);

		this.talker.Cast = this.cevioCast;
		this.talker.Volume = volume;
		this.talker.Speed = speed;
		this.talker.Tone = this.cevioTone;
		this.talker.ToneScale = this.cevioToneScale;
		this.talker.Alpha = this.cevioAlpha;
		this.talker.Volume = this.cevioVolume;
		try {
			var writtenWavLength = 0UL;
			var currentTextList = pTextFragList;
			while(true) {
				if(currentTextList.State.eAction == SPVACTIONS.SPVA_ParseUnknownTag) {
					goto next;
				}
				var text = Regex.Replace(
					currentTextList.pTextStart,
					@"<.+?>",
					"",
					RegexOptions.IgnoreCase);
				if(string.IsNullOrWhiteSpace(text)) {
					goto next;
				}
				if(((SPVESACTIONS)pOutputSite.GetActions()).HasFlag(SPVESACTIONS.SPVES_ABORT)) {
					return;
				}
				if(this.convertKana == "1") {
					text = English2Kana.Convert(text);
				}
				AddEventToSAPI(pOutputSite, currentTextList.pTextStart, text, writtenWavLength);

				var tmpWaveFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.wav");
				if(!deleteTmp(tmpWaveFile)) {
					goto next;
				}
				writtenWavLength += Speak(
					text,
					tmpWaveFile,
					pOutputSite,
					play,
					output);
				deleteTmp(tmpWaveFile);
			next:
				if(currentTextList.pNext == IntPtr.Zero) {
					break;
				} else {
					currentTextList = Marshal.PtrToStructure<SPVTEXTFRAG>(currentTextList.pNext);
				}
			}
		}
		catch {
			play($"{typeof(CeVioKanaTTSEngine).Namespace}.Resources.unknown-error.wav");
			throw;
		}
	}


	private uint Speak(
		string text,
		string wavPath,
		ISpTTSEngineSite pOutputSite,
		Action<string> play,
		Func<ISpTTSEngineSite, byte[], uint> output) {

		System.Diagnostics.Debug.Assert(this.cevio is not null);
		System.Diagnostics.Debug.Assert(this.talker is not null);

		var hFile = CreateFile(
			wavPath,
			GENERIC_READ,
			FILE_SHARE_READ | FILE_SHARE_WRITE,
			IntPtr.Zero,
			CREATE_NEW,
			0,
			IntPtr.Zero);
		try {
			if(!this.talker.OutputWaveToFile(text, wavPath)) {
				play($"{typeof(CeVioKanaTTSEngine).Namespace}.Resources.unknown-error.wav");
				return output(pOutputSite, new byte[4]);
			} else {
				using var ms = new MemoryStream();
				var b = new byte[76800]; // 1秒間のデータサイズ
				var pos = 0;
				var len = GetFileSize(hFile, IntPtr.Zero);
				while(ReadFile(
					hFile,
					b, b.Length,
					out var ret,
					IntPtr.Zero)) {

					if(0 < ret) {
						if(pos == 0) {
							var head = 104; // データ領域開始アドレス
							ms.Write(b, head, ret - head);
						} else {
							ms.Write(b, 0, ret);
						}
						pos += ret;
					}
					if(ret == 0 && len <= pos) {
						break;
					}
				}
				return output(pOutputSite, ms.ToArray());
			}
		}
		catch(Exception) {
			throw;
		}
		finally {
			CloseHandle(hFile);
		}
	}


	private void AddEventToSAPI(ISpTTSEngineSite outputSite, string allText, string speakTargetText, ulong writtenWavLength) {
		outputSite.GetEventInterest(out var ev);
		var list = new List<SPEVENT>();
		var wParam = (uint)speakTargetText.Length;
		var lParam = allText.IndexOf(speakTargetText);
		if((ev & SPFEI(SPEVENTENUM.SPEI_SENTENCE_BOUNDARY)) == SPFEI(SPEVENTENUM.SPEI_SENTENCE_BOUNDARY)) {
			list.Add(new SPEVENT() {
				eEventId = (ushort)SPEVENTENUM.SPEI_SENTENCE_BOUNDARY,
				elParamType = (ushort)SPEVENTLPARAMTYPE.SPET_LPARAM_IS_UNDEFINED,
				wParam = wParam,
				lParam = lParam,
				ullAudioStreamOffset = writtenWavLength
			});
		}
		if((ev & SPFEI(SPEVENTENUM.SPEI_WORD_BOUNDARY)) == SPFEI(SPEVENTENUM.SPEI_WORD_BOUNDARY)) {
			list.Add(new SPEVENT() {
				eEventId = (ushort)SPEVENTENUM.SPEI_WORD_BOUNDARY,
				elParamType = (ushort)SPEVENTLPARAMTYPE.SPET_LPARAM_IS_UNDEFINED,
				wParam = wParam,
				lParam = lParam,
				ullAudioStreamOffset = writtenWavLength
			});
		}
		if(list.Any()) {
			var arr = list.ToArray();
			outputSite.AddEvents(ref arr[0], (uint)arr.Length);
		}
	}

	public void GetOutputFormat(ref Guid pTargetFmtId, ref WAVEFORMATEX pTargetWaveFormatEx, out Guid pOutputFormatId, IntPtr ppCoMemOutputWaveFormatEx) {
		pOutputFormatId = SPDFID_WaveFormatEx;
		var wf = new WAVEFORMATEX() {
			wFormatTag = WAVE_FORMAT_PCM,
			nChannels = 1,
			cbSize = 0,
			nSamplesPerSec = 48000,
			wBitsPerSample = 16,
			nBlockAlign = 1 * 16 / 8, // チャンネル * bps / 8
			nAvgBytesPerSec = 48000 * (1 * 16 / 8), // サンプリングレート / ブロックアライン
		};

		var p = Marshal.AllocCoTaskMem(Marshal.SizeOf(wf));
		Marshal.StructureToPtr(wf, p, false);
		Marshal.WriteIntPtr(ppCoMemOutputWaveFormatEx, p);
	}

	public void SetObjectToken(ISpObjectToken pToken) {
		string get(string key) {
			try {
				pToken.GetStringValue(key, out var s);
				return s;
			}
			catch(COMException) {
				return "";
			}
		}
		uint @uint(string val, uint @default) {
			try {
				return uint.Parse(val);
			}
			catch {
				return @default;
			}
		}

		var ctlType = Type.GetTypeFromProgID("CeVIO.Talk.RemoteService.ServiceControlV40");
		var tlkType = Type.GetTypeFromProgID("CeVIO.Talk.RemoteService.TalkerV40");
		if((ctlType == null) || (tlkType == null)) {
			throw new Exception();
		}
		this.cevio = Activator.CreateInstance(ctlType);
		this.talker = Activator.CreateInstance(tlkType);
		if(this.cevio == null) {
			throw new Exception("CeVIOが見つかりません");
		}
		if(this.talker == null) {
			throw new Exception("CeVIOが見つかりません");
		}

		var start = (bool)this.cevio.IsHostStarted;
		if(!start) {
			this.cevio.StartHost(false);
		}

		this.token = pToken;
		this.cevioCast = get(KeyCeVioCast);
		this.cevioVolume = @uint(get(KeyCeVioVolume), DefaultVolume);
		this.cevioSpeed = @uint(get(KeyCeVioSpeed), DefaultSpeed);
		this.cevioTone = @uint(get(KeyCeVioTone), DefaultTone);
		this.cevioToneScale = @uint(get(KeyCeVioToneScale), DefaultToneScale);
		this.cevioAlpha = @uint(get(KeyCeVioAlpha), DefaultAlpha);
		var comp =get(KeyCeVioComponents);
		this.convertKana = get(KeyConvertKana);

		this.cevioComponents = comp.Split(',')
			.Select<string, (string, uint)?>(x => {
				var s = x.Split(':');
				if(s.Length == 2) {
					if(uint.TryParse(s[1].Trim(), out var val)) {
						return (s[0].Trim(), val);
					}
				}
				return null;
			}).Where(x => x != null)
			.Select<(string, uint)?, (string, uint)>(x => {
				System.Diagnostics.Debug.Assert(x != null);
				return x.Value;
			}).ToList().AsReadOnly();



		if(this.convertKana == "1" && !English2Kana.IsInited) {
			var dir = Path.GetDirectoryName(typeof(CeVioKanaTTSEngine).Assembly.Location);
			Lucene.Net.Configuration.ConfigurationSettings.GetConfigurationFactory().GetConfiguration()["kuromoji:data:dir"] = dir;
			English2Kana.Init();
		}
	}


	public void GetObjectToken(ref ISpObjectToken? ppToken) {
		ppToken = token;
	}

	[ComRegisterFunction()]
	public static void RegisterClass(string _) {
		static string safePath(string name) => Regex.Replace(name, @"[\s,/\:\*\?""\<\>\|]", "");
		var entry = @"SOFTWARE\Microsoft\Speech\Voices\Tokens";
		var prefix = "TTS_YARUKIZERO_CEVIO";


		// 一度情報を破棄する
		InitializeRegistry();
		var ctlType = Type.GetTypeFromProgID("CeVIO.Talk.RemoteService.ServiceControlV40");
		var tlkType = Type.GetTypeFromProgID("CeVIO.Talk.RemoteService.TalkerV40");
		if((ctlType == null) || (tlkType == null)) {
			throw new Exception();
		}
		dynamic? cevio = Activator.CreateInstance(ctlType);
		dynamic? tlk = Activator.CreateInstance(tlkType);
		if(cevio == null) {
			throw new Exception("CeVIOが見つかりません");
		}
		if(tlk == null) {
			throw new Exception("CeVIOが見つかりません");
		}

		var start = (bool)cevio.IsHostStarted;
		if(!start) {
			cevio.StartHost(false);
		}

		var regEntry = new List<(string Cast, string Components)>();
		var casts = tlk.AvailableCasts;
		// こっちはforeach使えない
		for(var i = 0; i < casts.Length; i++) {
			var components = new List<string>();
			var c = (string)casts.At(i);
			tlk.Cast = c;
			// こっちはforeachじゃないとだめ
			foreach(var comp in tlk.Components) {
				components.Add($"{comp.Name}:{comp.Value}");
			}
			regEntry.Add((c, string.Join(",", components)));
		}

		foreach(var it in regEntry) {
			using(var registryKey = Registry.LocalMachine.CreateSubKey($@"{entry}\{prefix}-{safePath(it.Cast)}")) {
				registryKey.SetValue("", $"CeVIO-Kana {it.Cast}");
				registryKey.SetValue("411", $"CeVIO-Kana {it.Cast}");
				registryKey.SetValue("CLSID", $"{{{GuidConst.ClassGuid}}}");
				registryKey.SetValue(KeyCeVioCast, it.Cast);
				registryKey.SetValue(KeyCeVioVolume, "50");
				registryKey.SetValue(KeyCeVioSpeed, "50");
				registryKey.SetValue(KeyCeVioTone, "50");
				registryKey.SetValue(KeyCeVioToneScale, "50");
				registryKey.SetValue(KeyCeVioAlpha, "50");
				registryKey.SetValue(KeyCeVioComponents, it.Components);
				registryKey.SetValue(KeyConvertKana, $"{1}");
			}
			using(var registryKey = Registry.LocalMachine.CreateSubKey($@"{entry}\{prefix}-{safePath(it.Cast)}\Attributes")) {
				registryKey.SetValue("Age", "Teen"); // ここはてきとー
				registryKey.SetValue("Vendor", "Hiroshiba Kazuyuki");
				registryKey.SetValue("Language", "411");
				registryKey.SetValue("Gender", "Female"); // ここもてきとー
				registryKey.SetValue("Name", $"CeVIO-Kana {it.Cast}");
			}
		}

		cevio.CloseHost(0);
	}
	

	[ComUnregisterFunction()]
	public static void UnregisterClass(string _) {
		InitializeRegistry();
	}

	private static void InitializeRegistry() {
		using(var regTokensKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Speech\Voices\Tokens\", true)) {
			if(regTokensKey == null) {
				return;
			}
			foreach(var name in regTokensKey.GetSubKeyNames()) {
				using(var regKey = regTokensKey.OpenSubKey(name)) {
					if(regKey?.GetValue("CLSID") is string id && id == $"{{{GuidConst.ClassGuid}}}") {
						regTokensKey.DeleteSubKeyTree(name);
					}
				}
			}
		}
	}
}