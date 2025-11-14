'''
Python Script
VoiceInserter
DaVinci Resolve向けのスクリプトで、Voicevoxなどの音声・画像・字幕を挿入するGUIを作成する。
'''
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import tkinter.ttk as ttk
import json
import re
import os
import sys
import glob
import wave
from io import BytesIO
from typing import Literal, Callable, Any, Final, cast
from uuid import UUID
import urllib.request
import urllib.error
import subprocess
import pprint

TRACK_TYPE_VIDEO_STRING: Final = "video"
TRACK_TYPE_AUDIO_STRING: Final = "audio"
TRACK_TYPE_SUBTITLE_STRING: Final = "subtitle"
TRACK_TYPES: Final = [TRACK_TYPE_VIDEO_STRING, TRACK_TYPE_AUDIO_STRING, TRACK_TYPE_SUBTITLE_STRING]
TRACK_TYPE_VIDEO: Final = 1
TRACK_TYPE_AUDIO: Final = 2
TRACK_TYPE_SUBTITLE: Final = 3
CLIP_NAME_PREFIX: Final = "VoiceInserter"
DATA_FILE: Final = "VoiceInserterData"
FONT_PATH: Final = "C:\\Windows\\Fonts"
scriptVersion: str = "1.0.0"
IGNORE_VERSION_FILE: Final = f"{os.environ['RESOLVE_SCRIPT_API']}/{DATA_FILE}/ignoreVersion.txt"

try:
    sys.path.append(f"{os.environ['RESOLVE_SCRIPT_API']}/Modules/voicevox_core/Lib/site-packages")
    import voicevox_core as voicevox
    from voicevox_core.blocking import Onnxruntime, OpenJtalk, Synthesizer, VoiceModelFile, UserDict
    import winsound
    import tempfile
    VOICEVOX_PATH: Final = f"{os.environ['RESOLVE_SCRIPT_API']}/../../../Fusion/Scripts/Utility/VoiceInserter/voicevox_core"
    VOICEVOX_TARGET_VERSION: Final = "0.16.2"
    voicevoxAvailable: bool = True
except:
    voicevoxAvailable = False

def GetWavDuration(wavedata: wave.Wave_read) -> float:
    '''
    waveデータの再生に掛かる秒数を取得する。
    事前にMakeWavを呼ぶ必要あり。

    Parameters:
    wavedata: wave.Wave_read
        waveのデータ

    Returns: float
        再生に掛かる秒数
    '''
    if not wavedata:
        return 0
    framerate: int = wavedata.getframerate()
    framecount: int = wavedata.getnframes()
    return framecount / framerate

def GetColorCode(r: float, g:float, b:float) -> str:
    '''
    カラーコードを取得する。
    
    Parameters:
    r: float
        赤(0-1)
    g: float
        緑(0-1)
    b: float
        青(0-1)
    
    Returns:str
        カラーコード
    '''
    return f"#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}"

class FontList:
    class FontStyles:
        def __init__(self, font) -> None:
            self.font: str = font
            self.styleSet: set[tuple[str, str]]= set()
            self.styleList: list[tuple[str, str]] = []
    
    @staticmethod
    def FetchFonts() -> dict[str, FontStyles]:
        '''
        インストール済みフォントのリストを返す.
        .ttfまたは.ttcのみ扱う

        Returns: dict[str, tuple[str, set[tuple[str, str]]]]
            フォントとスタイルのセット
        '''
        JAPANESE_NAME: Final = 0
        ENGLISH_NAME: Final = 1
        FONT_ID: Final = (1, 16)
        STYLE_ID: Final = (2, 17)
        ttfEndian: Literal['little', 'big'] = "big"
        retFonts: dict[str, FontList.FontStyles] = {}
        for fileName in glob.glob(f"{FONT_PATH}\\*"):
            if not (fileName.upper().endswith('.TTF') or fileName.upper().endswith('.TTC')):
                continue
            with open(fileName, "rb") as f:
                # nameテーブルの検索
                nameOffset: int | None = None
                version: int = int.from_bytes(f.read(4), ttfEndian)
                if version == 0x00010000:
                    offsetTable: list[int] = [0x0]
                elif version == 0x74746366: # ttcf
                    offsetTable = []
                    f.seek(4, 1)
                    numFonts: int = int.from_bytes(f.read(4), ttfEndian)
                    for i in range(numFonts):
                        offsetTable.append(int.from_bytes(f.read(4), ttfEndian))
                else:
                    print(f"Unknown font file version: {version} in {fileName}")
                    continue
                for offset in offsetTable:
                    f.seek(offset, 0)
                    version = int.from_bytes(f.read(4), ttfEndian) 
                    if version != 0x00010000:
                        print(f"Unknown font sub file version: {version} in {fileName}")
                        continue
                    tableCount: int = int.from_bytes(f.read(2), ttfEndian)
                    f.seek(6, 1)
                    for i in range(tableCount):
                        try:
                            tagName: str = f.read(4).decode("ascii")
                        except UnicodeDecodeError as e:
                            print(f"{fileName}: UnicodeDecodeError in table tag name pos: 0x{f.tell() - 4:x}")
                            raise e

                        if tagName == 'name':
                            f.seek(4, 1)
                            nameOffset = int.from_bytes(f.read(4), ttfEndian)
                            break
                        f.seek(12, 1)
                    if nameOffset is None:
                        continue
                    f.seek(nameOffset, 0)
                    # フォント名とスタイル名の検索
                    f.seek(2, 1)
                    nameRecordCount: int = int.from_bytes(f.read(2), ttfEndian)
                    storageOffset: int = int.from_bytes(f.read(2), ttfEndian)
                    nameRecordOffset: int = f.tell()
                    nameRecordLength: int = 12

                    fonts: dict[int, list[str]] = {}
                    styles: dict[int, list[str]] = {}
                    for i in range(nameRecordCount):
                        f.seek(nameRecordOffset + nameRecordLength * i, 0)
                        platformID: int = int.from_bytes(f.read(2), ttfEndian)
                        encodingID: int = int.from_bytes(f.read(2), ttfEndian)
                        languageID: int = int.from_bytes(f.read(2), ttfEndian)
                        nameID: int = int.from_bytes(f.read(2), ttfEndian)
                        length: int = int.from_bytes(f.read(2), ttfEndian)
                        stringOffset: int = int.from_bytes(f.read(2), ttfEndian)
                        if not nameID in FONT_ID and not nameID in STYLE_ID:
                            continue
                        f.seek(nameOffset + storageOffset + stringOffset, 0)
                        data = f.read(length)
                        if platformID == 0 or (platformID == 3 and encodingID == 1):
                            string: str = data.decode("utf-16-be")
                        elif platformID == 1 and encodingID == 0:
                            string = data.decode("mac_roman")
                        elif platformID == 1 and encodingID == 1 or platformID == 3 and encodingID == 2:
                            string = data.decode("shift_jis")
                        else:
                            continue
                        if platformID == 0:
                            japaneseLanguageIDList: list[int] = [0]
                            englishLanguageIDList: list[int] = [0]
                        elif platformID == 1:
                            japaneseLanguageIDList = [11]
                            englishLanguageIDList = [0]
                        elif platformID == 3:
                            japaneseLanguageIDList = [0x0411, 0x0011, 0x0811]
                            englishLanguageIDList = [0x1000, 0x0409]
                        else:
                            continue
                        if nameID in FONT_ID:
                            if nameID not in fonts:
                                fonts[nameID] = ["", ""]
                            # 日本語
                            if languageID in japaneseLanguageIDList:
                                fonts[nameID][JAPANESE_NAME] = string
                            # 英語.
                            if languageID in englishLanguageIDList:
                                fonts[nameID][ENGLISH_NAME] = string
                        elif nameID in STYLE_ID:
                            if nameID not in styles:
                                styles[nameID] = ["", ""]
                            # 日本語
                            if languageID in japaneseLanguageIDList:
                                styles[nameID][JAPANESE_NAME] = string
                            # 英語.
                            if languageID in englishLanguageIDList:
                                styles[nameID][ENGLISH_NAME] = string
                    if len(fonts) > 0:
                        if FONT_ID[1] in fonts:
                            font: str = fonts[FONT_ID[1]][ENGLISH_NAME] if fonts[FONT_ID[1]][ENGLISH_NAME] != "" else fonts[FONT_ID[1]][JAPANESE_NAME]
                            dispFont: str = fonts[FONT_ID[1]][JAPANESE_NAME] if fonts[FONT_ID[1]][JAPANESE_NAME] != "" else fonts[FONT_ID[1]][ENGLISH_NAME]
                        else:
                            font = fonts[FONT_ID[0]][ENGLISH_NAME] if fonts[FONT_ID[0]][ENGLISH_NAME] != "" else fonts[FONT_ID[0]][JAPANESE_NAME]
                            dispFont = fonts[FONT_ID[0]][JAPANESE_NAME] if fonts[FONT_ID[0]][JAPANESE_NAME] != "" else fonts[FONT_ID[0]][ENGLISH_NAME]
                        if not dispFont in retFonts:
                            retFonts[dispFont] = FontList.FontStyles(font)
                        if len(styles) > 0:
                            if STYLE_ID[1] in styles:
                                style: str = styles[STYLE_ID[1]][ENGLISH_NAME] if styles[STYLE_ID[1]][ENGLISH_NAME] != "" else styles[STYLE_ID[1]][JAPANESE_NAME]
                                dispStyle: str = styles[STYLE_ID[1]][JAPANESE_NAME] if styles[STYLE_ID[1]][JAPANESE_NAME] != "" else styles[STYLE_ID[1]][ENGLISH_NAME]
                            else:
                                style = styles[STYLE_ID[0]][ENGLISH_NAME] if styles[STYLE_ID[0]][ENGLISH_NAME] != "" else styles[STYLE_ID[0]][JAPANESE_NAME]
                                dispStyle = styles[STYLE_ID[0]][JAPANESE_NAME] if styles[STYLE_ID[0]][JAPANESE_NAME] != "" else styles[STYLE_ID[0]][ENGLISH_NAME]
                            retFonts[dispFont].styleSet.add((dispStyle, style))
        for dispFont in retFonts.keys():
            retFonts[dispFont].styleList = list(retFonts[dispFont].styleSet)
        return retFonts

    def __init__(self, fontDict: dict[str, FontStyles]) -> None:
        self.dispFonts = fontDict.keys()
        self.fonts: dict[str, str] = {}
        for font in self.dispFonts:
            self.fonts[font] = fontDict[font].font
        self.dispStyles: dict[str, list[str]] = {}
        self.style: dict[str, dict[str, str]] = {}
        for font in self.fonts:
            self.dispStyles[font] = []
            self.style[font] = {}
            for style in fontDict[font].styleList:
                self.dispStyles[font].append(style[0])
                self.style[font][style[0]] = style[1]

class ResolveUtil:
    @staticmethod
    def TimecodeToFrames(timecode: str, fps: int | float) -> int:
        '''
        タイムコードをフレーム数に変換する
        Parameters:
        timecode: string
            タイムコード (HH:MM:SS:FF)
        fps: int
            フレームレート
        Returns: int
            フレーム数
        '''
        h, m, s, f = re.findall(r"\d+", timecode)
        return int((int(h) * 3600 + int(m) * 60 + int(s)) * fps) + int(f)

    @staticmethod
    def GetTimecodeFromFrame(frame: int, fps: int | float) -> str:
        '''
        フレームをタイムコードに変換する。

        Parameters:
        frame: int
            フレーム数
        fps: int
            フレームレート

        Returns: string
            タイムコード(HH:MM:SS:FF)
        '''
        return ResolveUtil.AddFrameToTimecode("00:00:00:00", frame, fps)
    
    @staticmethod
    def AddFrameToTimecode(timecode: str, addFrames: int, fps: int | float) -> str:
        '''
        タイムコードにフレーム数を加算する

        Parameters:
        timecode: string
            タイムコード (HH:MM:SS:FF)
        addFrames: int
            加算するフレーム数
        fps: int
            フレームレート
        Returns: string
            加算後のタイムコード (HH:MM:SS:FF)
        '''
        totalFrames = ResolveUtil.TimecodeToFrames(timecode, fps) + addFrames
        if totalFrames < 0:
            totalFrames = 0
        hours = int(totalFrames // (3600 * fps))
        minutes = int((totalFrames % (3600 * fps)) // (60 * fps))
        seconds = int((totalFrames % (60 * fps)) // fps)
        frames = int(totalFrames % fps)
        return f"{hours:02}:{minutes:02}:{seconds:02}:{frames:02}"

    @staticmethod
    def MoveCurrentFolder(mediaPool, folderPath: str, createIfNotExist: bool=True) -> None:
        '''
        メディアプールのフォルダを移動する

        Parameters:
        mediaPool: mediaPool
            操作するメディアプール
        folderPath: string
            フォルダのパス。/で区切る
        createIfNotExist: bool
            フォルダが存在しない場合に新規作成するかどうか
        '''
        if not mediaPool:
            messagebox.showerror("Error", "有効なメディアプールがありません。")
            return None
        # ルートフォルダからの絶対パスの場合はひとまずルートフォルダに移動する
        if folderPath[0] == '/':
            mediaPool.SetCurrentFolder(mediaPool.GetRootFolder())
        currentFolder = mediaPool.GetCurrentFolder()
        if not folderPath:
            return currentFolder
        folderNames = folderPath.split('/')
        for folderName in folderNames:
            if not folderName:
                continue
            subFolders = currentFolder.GetSubFolders()
            for subFolder in subFolders.values():
                if subFolder.GetName() == folderName:
                    currentFolder = subFolder
                    break
            else:
                if not createIfNotExist:
                    messagebox.showerror("Error", f"フォルダ '{folderName}' が存在しません。")
                    return 
                newFolder = mediaPool.AddSubFolder(currentFolder, folderName)
                if not newFolder:
                    messagebox.showerror("Error", f"フォルダ '{folderName}' の作成に失敗しました。")
                    return 
                currentFolder = newFolder
            mediaPool.SetCurrentFolder(currentFolder)

    @staticmethod
    def GetOrCreateCurrentTimeline(project):
        '''
        プロジェクトの現在のタイムラインを取得する
        存在しない場合は新規に作成する

        Parameters:
        project: project
            操作するプロジェクト

        Returns: timeline
            取得したタイムライン
        '''
        if not project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return None
        currentTimeline = project.GetCurrentTimeline()
        if not currentTimeline:
            mediaPool = project.GetMediaPool()
            currentTimeline = mediaPool.CreateEmptyTimeline("Timeline 1")
            project.SetCurrentTimeline(currentTimeline)
        return currentTimeline

    @staticmethod
    def SearchTrackIndex(timeline, trackType: Literal["video", "audio", "subtitle"], trackName: str) -> int:
        '''
        トラック名からトラックのインデックスを返す

        Parameters:
        timeline: timeline
            探すタイムライン
        trackType: TRACK_TYPES
            探すトラックのタイプ
        trackName: str
            探すトラック名

        Returns:int
            トラックのインデックス。見つからなければ-1
        '''
        if not timeline:
            return -1
        trackCount = timeline.GetTrackCount(trackType) + 1
        for i in range(1, trackCount):
            if timeline.GetTrackName(trackType, i) == trackName:
                return i
        return -1

    @staticmethod
    def GetCurrentTimelineClip(project, trackType: Literal["video", "audio", "subtitle"], trackName: str):
        '''
        現在のタイムラインの指定されたトラック名のクリップを取得する

        Parameters:
        project: project
            操作するプロジェクト
        trackType: "video" or "audio"
            トラックの種類
        trackName: string
            トラック名

        Returns: timelineClip
            取得したクリップ
        '''
        if not project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return None
        currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(project)
        trackIndex = ResolveUtil.SearchTrackIndex(currentTimeline, trackType, trackName)
        if trackIndex == -1:
            messagebox.showerror("Error", f"{trackType}トラック '{trackName}' が見つかりませんでした。")
            return None
        currentTime = currentTimeline.GetCurrentTimecode()
        fps = currentTimeline.GetSetting("timelineFrameRate")
        if currentTime is None:
            messagebox.showerror("Error", "タイムラインが表示された画面ではありません。")
            return None
        currentFrame = ResolveUtil.TimecodeToFrames(currentTime, fps)
        clips = currentTimeline.GetItemsInTrack(trackType, trackIndex)
        for clip in clips.values():
            if clip.GetStart(False) <= currentFrame <= clip.GetEnd(False):
                return clip
        return None

class TkinterUtil:
    class SubWindow(tk.Toplevel):
        def __init__(self, master: tk.Misc | None, OnDestroy: Callable[[], Any] | None = None):
            super().__init__(master)
            self.protocol('WM_DELETE_WINDOW', self.destroy)
            self.OnDestroy: Callable[[], Any] | None = OnDestroy
        
        def destroy(self):
            super().destroy()
            if self.OnDestroy is not None:
                self.OnDestroy()

class VoicevoxEngine:
    class ModelInfo:
        def __init__(self, file: str, id: int) -> None:
            self.filename: str = file
            self.styleID: int = id

    def __init__(self) -> None:
        self.__text: str = ""
        self.__currentStyleID: int = -1
        self.__wav: bytes | None = None
        self.__loadedModel: voicevox.blocking.VoiceModelFile | None = None
        self.__temp: tempfile._TemporaryFileWrapper | None = None
        self.__voiceModelList: dict[str, dict[str, VoicevoxEngine.ModelInfo]] = {}
        self._userDict: voicevox.blocking.UserDict | None = None
        self._accentPhrases: list[voicevox.AccentPhrase] | None = None
        self._intonationFrame: tk.Frame | None = None
        self._moraLengthFrame: tk.Frame | None = None
        self._listboxFrame: tk.Frame | None = None
        self._DictionaryEditFrame: tk.Frame | None = None
        self._user_dict_path: str = f"{VOICEVOX_PATH}/dict/user.dic"
        voicevox_onnxruntime_path: str = f"{VOICEVOX_PATH}/onnxruntime/lib/{Onnxruntime.LIB_VERSIONED_FILENAME}" # type: ignore[attr-defined]
        open_jtalk_dict_dir: str = f"{VOICEVOX_PATH}/dict/open_jtalk_dic_utf_8-1.11"
        #OpenJTalkの初期化
        self.__open_jtalk: voicevox.blocking.OpenJtalk = OpenJtalk(open_jtalk_dict_dir)
        if not self.__open_jtalk:
            # 失敗
            print(f"open jtalk rc new error")
            return
        #Synthesizerの初期化
        self.__synthesizer: voicevox.blocking.Synthesizer = Synthesizer(Onnxruntime.load_once(filename=voicevox_onnxruntime_path), self.__open_jtalk) # type: ignore[attr-defined]
        if not self.__synthesizer:
            # 失敗
            print(f"synthesizer new error")
            return
        self._MakeVoiceModelList()
        self.LoadUserDict()
    
    def __del__(self) -> None:
        self.StopPlayWav()

    def IsInitSucceeded(self) -> bool:
        return self.__synthesizer is not None
    
    def GetCharacterList(self) -> list[str]:
        '''
        キャラのリストを取得

        Returns: List(str)
            キャラ名のリスト
        '''
        return list(self.__voiceModelList.keys())
    
    def GetStyleList(self, character) -> list[str]:
        '''
        指定したキャラのスタイルのリストを取得

        Parameters:
        character: str
            スタイルリストを取得するキャラクター名
        
        Returns: List(str)
            スタイル名のリスト
        '''
        return list(self.__voiceModelList[character].keys())

    def MakeVoice(self, charaname: str, stylename: str, text: str, upspeak: bool, speed: float=1.0, pitch: float=0.0, intonation: float=1.0, volume: float=1.0, pauseLengthScale: float=1.0, prePhonemeLength: float=0.1, postPhonemeLength: float=0.1) -> bool:
        '''
        ボイスのwavデータを作成する.

        paramters:
        charaname: str
            声のキャラ名
        stylename: str
            スタイル名
        text: str
            読ませるテキスト
        upspeak: bool
            疑問文の調整を有効にするか
        speed: float
            話速
        pitch: float
            音高
        intonation: float
            抑揚
        volume: float
            音量
        pauseLengthScale: float
            無音時間スケール
        prePhonemeLength: float
            前の無音長さ
        postPhonemeLength: float
            後の無音長さ

        returns: bool
            成否
        '''
        # Modelのロード
        self._LoadVoiceModel(charaname, stylename)
        styleID: int = self.__voiceModelList[charaname][stylename].styleID
        if self.__text != text or self.__currentStyleID != styleID:
            # 文章が違っていればアクセント句を取得
            self.__text = text
            self.__currentStyleID = styleID
            self._accentPhrases = self.__synthesizer.create_accent_phrases(text, self.__currentStyleID)
            self._UpdatePhraseEditorDisp()
        if self._accentPhrases is None:
            return False
        # pauseLengthScaleの適用.
        defaultPauseLengthes: list[float] = []
        for i in range(len(self._accentPhrases)):
            accentPhrase = self._accentPhrases[i]
            defaultPauseLengthes.append(0)
            if accentPhrase.pause_mora is not None:
                if accentPhrase.pause_mora.vowel == "pau":
                    defaultPauseLengthes[i] = accentPhrase.pause_mora.vowel_length
                    accentPhrase.pause_mora.vowel_length *= pauseLengthScale
        # AudioQueryのパラメータ設定
        audioQuery: voicevox.AudioQuery = voicevox.AudioQuery.from_accent_phrases(self._accentPhrases)
        audioQuery.speed_scale = speed
        audioQuery.pitch_scale = pitch
        audioQuery.intonation_scale = intonation
        audioQuery.volume_scale = volume
        audioQuery.pre_phoneme_length = prePhonemeLength
        audioQuery.post_phoneme_length = postPhonemeLength

        # wavの作成
        self.__wav = self.__synthesizer.synthesis(audioQuery, self.__voiceModelList[charaname][stylename].styleID, enable_interrogative_upspeak=upspeak)
        # pauseLengthScaleの適用を戻す
        for i in range(len(self._accentPhrases)):
            accentPhrase = self._accentPhrases[i]
            defaultPauseLengthes.append(0)
            if accentPhrase.pause_mora is not None:
                if accentPhrase.pause_mora.vowel == "pau":
                    accentPhrase.pause_mora.vowel_length = defaultPauseLengthes[i]
        if not self.__wav:
            return False
        return True
        
    def SaveWav(self, filepath: str) -> None:
        '''
        WAVデータを保存する。
        事前にMakeWavを呼ぶ必要あり

        Parameters:
        filepath: str
            保存先
        '''
        if not self.__wav:
            return
        # wavファイルに保存.
        with open(filepath, "wb") as f:
            f.write(self.__wav)
        return

    def PlayWav(self) -> None:
        '''
        waveデータを再生する。
        事前にMakeWavを呼ぶ必要あり
        '''
        if not self.__wav:
            return
        if self.__temp:
            self.StopPlayWav()
        self.__temp = tempfile.NamedTemporaryFile(delete=False)
        self.__temp.write(self.__wav)
        winsound.PlaySound(self.__temp.name, winsound.SND_FILENAME | winsound.SND_ASYNC)

    def CalcWavDuration(self) -> float:
        '''
        waveデータの再生に掛かる秒数を取得する。
        事前にMakeWavを呼ぶ必要あり。

        Returns: float
            再生に掛かる秒数
        '''
        if not self.__wav:
            return 0
        wavFile: BytesIO = BytesIO(self.__wav) 
        wavedata: wave.Wave_read = wave.open(wavFile)
        return GetWavDuration(wavedata)

    def StopPlayWav(self) -> None:
        '''
        再生中のwaveデータを止める。
        '''
        winsound.PlaySound(None, winsound.SND_FILENAME)
        if self.__temp:
            self.__temp.close()
            os.remove(self.__temp.name)
            self.__temp = None

    def MergeAccentPhrase(self, mergeIndex: int) -> Callable[[], None]:
        '''
        アクセントフレーズを合体させる関数を返す
        
        Parameters:
        mergeIndex: int
            合体させる手前のインデックス

        Returns: Func()
            アクセントフレーズを合体させる関数
        '''
        def inner() -> None:
            if self._accentPhrases is None or len(self._accentPhrases) <= mergeIndex + 1:
                return
            newMoras: list[voicevox.Mora] = self._accentPhrases[mergeIndex].moras + self._accentPhrases[mergeIndex+1].moras
            newAccent: int = self._accentPhrases[mergeIndex].accent
            newPauseMora: voicevox.Mora | None = self._accentPhrases[mergeIndex+1].pause_mora
            newIsInterrogative: bool = self._accentPhrases[mergeIndex+1].is_interrogative
            newAccentPhrase: voicevox.AccentPhrase = voicevox.AccentPhrase(newMoras, newAccent, newPauseMora, newIsInterrogative)
            self._accentPhrases[mergeIndex:mergeIndex+2] = [newAccentPhrase]
            self._UpdateMoraData()
            self._UpdatePhraseEditorDisp()
        return inner
    
    def SplitAccentPhrase(self, splitPhraseIndex: int, splitMoraIndex: int) -> Callable[[], None]:
        '''
        アクセントフレーズを指定モーラまでのものとに分割する関数を返す
        
        Parameters:
        splitIndex: int
            分割するアクセントフレーズ
        splitMoraIndex: int
            新しいアクセントフレーズの手前側のものの最後のモーラ
        
        Returns: Func()
            アクセントフレーズを分割する関数
        '''
        def inner() -> None:
            if self._accentPhrases is None or len(self._accentPhrases) <= splitPhraseIndex or len(self._accentPhrases[splitPhraseIndex].moras) <= splitMoraIndex - 1:
                return
            newMorasFormer: list[voicevox.Mora] = self._accentPhrases[splitPhraseIndex].moras[:splitMoraIndex + 1]
            newAccentFormer: int = self._accentPhrases[splitPhraseIndex].accent if splitMoraIndex >= self._accentPhrases[splitPhraseIndex].accent else splitMoraIndex + 1
            newPauseMoraFormer: voicevox.Mora | None = None
            newAccentPhraseFormer: voicevox.AccentPhrase = voicevox.AccentPhrase(newMorasFormer, newAccentFormer, newPauseMoraFormer)
            newMorasLatter: list[voicevox.Mora] = self._accentPhrases[splitPhraseIndex].moras[splitMoraIndex+1:]
            newAccentLatter: int = self._accentPhrases[splitPhraseIndex].accent - splitMoraIndex - 1 if splitMoraIndex + 1 < self._accentPhrases[splitPhraseIndex].accent else 1
            newPauseMoraLatter: voicevox.Mora | None = self._accentPhrases[splitPhraseIndex].pause_mora
            newAccentPhraseLatter: voicevox.AccentPhrase = voicevox.AccentPhrase(newMorasLatter, newAccentLatter, newPauseMoraLatter)
            self._accentPhrases[splitPhraseIndex:splitPhraseIndex+1] = [newAccentPhraseFormer, newAccentPhraseLatter]
            self._UpdateMoraData()
            self._UpdatePhraseEditorDisp()
        return inner

    def DeleteAccentPhrase(self, deleteIndex: int) -> Callable[[], None]:
        '''
        アクセントフレーズを削除する関数を返す

        Parameters:
        deleteIndex: int
            削除するアクセントフレーズ

        Returns: Func()
            アクセントフレーズを削除する関数
        '''
        def inner():
            if self._accentPhrases is None or len(self._accentPhrases) <= deleteIndex:
                return
            self._accentPhrases.pop(deleteIndex)
            self._UpdatePhraseEditorDisp()
        return inner

    def InitPhraseEditorDisp(self, frame: tk.Misc | None) -> None:
        '''
        アクセントフレーズ編集画面の初期化

        Parameters:
        frame: tkinter.frame
            親フレーム
        
        '''

        style: ttk.Style = ttk.Style()
        style.configure("VoicevoxUtil.TNotebook", tabposition="wn")
        phraseNote: ttk.Notebook = ttk.Notebook(frame, style="VoicevoxUtil.TNotebook")
        phraseNote.pack(fill='both', expand=True, padx=5)
        self._intonationFrame = tk.Frame(phraseNote)
        phraseNote.add(self._intonationFrame, text="イントネーション")
        self._moraLengthFrame = tk.Frame(phraseNote)
        phraseNote.add(self._moraLengthFrame, text="長さ")
        dictionaryEditButton: ttk.Button = ttk.Button(frame, text="辞書編集", command=self.OpenDictionaryEditor(frame))
        dictionaryEditButton.pack(side=tk.LEFT)

    def LoadUserDict(self) -> None:
        self._userDict = UserDict()
        if os.path.exists(self._user_dict_path):
            self._userDict.load(self._user_dict_path)
 
    def SearchUserDictWordUUID(self, userDictWord: voicevox.UserDictWord) -> UUID | None:
        if self._userDict is None:
            return None
        wordsDic = self._userDict.to_dict()
        for uuid, word in wordsDic.items():
            if word == userDictWord:
                return uuid
        return None
        
    def AddUserDict(self, userDictWord: voicevox.UserDictWord) -> None:
        if self._userDict is None:
            return
        self._userDict.add_word(userDictWord)
        self._UpdateUserDict()
        self._UpdateDictionaryEditorList(None)

    def UpdateUserDictWord(self, uuid: UUID | None, newUserDictWord: voicevox.UserDictWord) -> None:
        if self._userDict is None or uuid is None:
            return
        self._userDict.update_word(uuid, newUserDictWord)
        self._UpdateUserDict()
        self._UpdateDictionaryEditorEdit(None, newUserDictWord)
    
    def DelUserDictWord(self, userDictWord: voicevox.UserDictWord) -> None:
        if self._userDict is None:
            return
        uuid = self.SearchUserDictWordUUID(userDictWord)
        if uuid is not None:
            self._userDict.remove_word(uuid)
            self._UpdateUserDict()
            self._UpdateDictionaryEditorList(None)

    def DictionaryEditorDisp(self, windowRoot: tk.Misc) -> None:
        panedWindow: ttk.Panedwindow = ttk.Panedwindow(windowRoot, orient=tk.HORIZONTAL)
        leftFrame: tk.Frame = tk.Frame(panedWindow)
        self._UpdateDictionaryEditorList(leftFrame)

        panedWindow.add(leftFrame, weight=1) 

        rightFrame: tk.Frame = tk.Frame(panedWindow)
        self._UpdateDictionaryEditorEdit(rightFrame, None)
        panedWindow.add(rightFrame, weight=3)
        panedWindow.pack(fill=tk.BOTH, expand=True)

    def OpenDictionaryEditor(self, root) -> Callable[[], None]:
        def inner() -> None:
            def OnDestroy() -> None:
                self._accentPhrases = None
                self._UpdatePhraseEditorDisp()
                self._listboxFrame = None
                self._DictionaryEditFrame = None

            self._accentPhrases = None
            self._UpdatePhraseEditorDisp()
            editorRoot: TkinterUtil.SubWindow = TkinterUtil.SubWindow(root, OnDestroy)
            editorRoot.transient(root)
            editorRoot.title("ユーザー辞書")
            self.DictionaryEditorDisp(editorRoot)
            master: Any = root
            while master.master is not None:
                master = master.master
            editorRoot.grab_set()
            master.wait_window(editorRoot)
        return inner

    def _UpdatePhraseEditorDisp(self) -> None:
        '''
        アクセントフレーズの編集画面をtkinterで表示する。
        '''

        if self._intonationFrame is None or self._moraLengthFrame is None:
            print("アクセントフレーズ編集画面の初期化がされる前に表示に来た")
            return
        self._DeletePhraseEditor()
        if self._accentPhrases is None:
            return
        def GetScrollEvent(canvas: tk.Canvas) -> Callable[[tk.Event], None]:
            def inner(event: tk.Event):
                if event.delta > 0:
                    canvas.xview_scroll(-1, "units")
                elif event.delta < 0:
                    canvas.xview_scroll(1, "units")
            return inner
        # イントネーションタブ
        intonationCanvas: tk.Canvas = tk.Canvas(self._intonationFrame)
        intonationInnerFrame: tk.Frame = tk.Frame(intonationCanvas)
        intonationScrollbar: tk.Scrollbar = tk.Scrollbar(self._intonationFrame, orient=tk.HORIZONTAL, command=intonationCanvas.xview)
        self._intonationFrame.grid_columnconfigure(0, weight=1)
        intonationCanvas.grid(column=0, row=0, sticky=tk.W + tk.E)
        intonationScrollbar.grid(column=0, row=1, sticky=tk.W + tk.E)
        intonationCanvas.bind("<MouseWheel>", GetScrollEvent(intonationCanvas))
        intonationInnerFrame.bind("<MouseWheel>", GetScrollEvent(intonationCanvas))
        # 操作部
        def OnChangeIntonationScale(accentPhraseIndex: int, moraIndex: int) -> Callable[[str], None]:
            def inner(value: str) -> None:
                if self._accentPhrases is None:
                    return
                self._accentPhrases[accentPhraseIndex].moras[moraIndex].pitch = float(value)
            return inner
        def OnChangeAccentButton(accentPhraseIndex, value) -> Callable[[], None]:
            def inner() -> None:
                if self._accentPhrases is None:
                    return
                self._accentPhrases[accentPhraseIndex].accent = int(value.get())
                self._UpdateMoraData()
                self._UpdatePhraseEditorDisp()
            return inner
        # HACK:canvasのwidth/heightの取り方がよくない。
        canvasWidth: int = 0
        canvasHeight: int = 0
        for i in range(len(self._accentPhrases)):
            if i > 0:
                mergeButton: ttk.Button = ttk.Button(intonationInnerFrame, text="<>", width=3, command=self.MergeAccentPhrase(i-1))
                mergeButton.pack(side=tk.LEFT)
                mergeButton.update()
                canvasWidth += mergeButton.winfo_width() + 2
            accentPhrase: voicevox.AccentPhrase = self._accentPhrases[i]
            accentPhraseFrame: tk.Frame = tk.Frame(intonationInnerFrame, relief=tk.RAISED, bd=2)
            accentPhraseFrame.bind("<MouseWheel>", GetScrollEvent(intonationCanvas))
            accentPhraseFrame.pack(side=tk.LEFT)
            canvasWidth += 16
            columnIndex: int = 0
            accentValue: tk.StringVar = tk.StringVar(accentPhraseFrame, value=f"{accentPhrase.accent}")
            accentHeight: int = 0
            for j in range(len(accentPhrase.moras)):
                mora: voicevox.Mora = accentPhrase.moras[j]
                moraHeight: int = 0
                if j > 0:
                    splitButton: ttk.Button = ttk.Button(accentPhraseFrame, text="><", width=3, command=self.SplitAccentPhrase(i, j-1))
                    splitButton.grid(column=columnIndex, row=0, rowspan=3)
                    splitButton.update()
                    canvasWidth += splitButton.winfo_width() + 2
                    columnIndex += 1
                accentRadioButton: ttk.Radiobutton = ttk.Radiobutton(accentPhraseFrame, variable=accentValue, value=f"{j+1}", command=OnChangeAccentButton(i, accentValue))
                accentRadioButton.grid(column=columnIndex, row=0)
                accentRadioButton.update()
                moraHeight += accentRadioButton.winfo_height()
                pitchScale: tk.Scale = tk.Scale(accentPhraseFrame, from_=6.50, to=3.00, resolution=0.01, orient=tk.VERTICAL, command=OnChangeIntonationScale(i, j))
                pitchScale.set(mora.pitch)
                pitchScale.grid(column=columnIndex, row=1)
                pitchScale.update()
                canvasWidth += pitchScale.winfo_width() + 2
                moraHeight += pitchScale.winfo_height()
                textLabel: tk.Label = tk.Label(accentPhraseFrame, text=mora.text)
                textLabel.grid(column=columnIndex, row=2)
                textLabel.update()
                moraHeight += textLabel.winfo_height()
                if moraHeight > accentHeight:
                    accentHeight = moraHeight
                columnIndex += 1
            deleteButton: ttk.Button = ttk.Button(accentPhraseFrame, text="削除", command=self.DeleteAccentPhrase(i))
            deleteButton.grid(column=0, row=3, columnspan=columnIndex)
            deleteButton.update()
            accentHeight += deleteButton.winfo_height()
            if accentHeight > canvasHeight:
                canvasHeight = accentHeight
        canvasHeight += 2
        intonationCanvas.configure(scrollregion=(0, 0, canvasWidth, canvasHeight))
        intonationCanvas.configure(xscrollcommand=intonationScrollbar.set)
        intonationCanvas.create_window((0, 0), window=intonationInnerFrame, anchor="nw", width=canvasWidth, height=canvasHeight)

        # モーラ長さタブ
        moraLengthCanvas: tk.Canvas = tk.Canvas(self._moraLengthFrame)
        moraLengthInnerFrame: tk.Frame = tk.Frame(moraLengthCanvas)
        moraLengthScrollbar: tk.Scrollbar = tk.Scrollbar(self._moraLengthFrame, orient=tk.HORIZONTAL, command=moraLengthCanvas.xview)
        self._moraLengthFrame.grid_columnconfigure(0, weight=1)
        moraLengthCanvas.grid(column=0, row=0, sticky=tk.W + tk.E)
        moraLengthScrollbar.grid(column=0, row=1, sticky=tk.W + tk.E)
        moraLengthCanvas.bind("<MouseWheel>", GetScrollEvent(moraLengthCanvas))
        moraLengthInnerFrame.bind("<MouseWheel>", GetScrollEvent(moraLengthCanvas))
        # 操作部
        # HACK:canvasのwidth/heightの取り方がよくない。
        canvasWidth = 0
        canvasHeight = 0
        def OnChangeMoraLengthScale(accentPhraseIndex: int, moraIndex: int, isVowel: bool) -> Callable[[str], None]:
            def inner(value: str) -> None:
                if self._accentPhrases is None:
                    return
                if isVowel:
                    self._accentPhrases[accentPhraseIndex].moras[moraIndex].vowel_length = float(value)
                else:
                    self._accentPhrases[accentPhraseIndex].moras[moraIndex].consonant_length = float(value)
            return inner
        for i in range(len(self._accentPhrases)):
            if i > 0:
                mergeButton = ttk.Button(moraLengthInnerFrame, text="<>", width=3, command=self.MergeAccentPhrase(i-1))
                mergeButton.pack(side=tk.LEFT)
                mergeButton.update()
                canvasWidth += mergeButton.winfo_width() + 2
            accentPhrase = self._accentPhrases[i]
            accentPhraseFrame = tk.Frame(moraLengthInnerFrame)
            accentPhraseFrame.bind("<MouseWheel>", GetScrollEvent(moraLengthCanvas))
            accentPhraseFrame.pack(side=tk.LEFT)
            canvasWidth += 16
            maxMoraHeight = 0
            columnIndex = 0
            for j in range(len(accentPhrase.moras)):
                mora = accentPhrase.moras[j]
                moraHeight = 0
                if mora.consonant:
                    consonantLengthScale: tk.Scale = tk.Scale(accentPhraseFrame, from_=0.3, to=0, resolution=0.01, orient=tk.VERTICAL, command=OnChangeMoraLengthScale(i, j, False))
                    consonantLengthScale.set(mora.consonant_length)
                    consonantLengthScale.grid(column=columnIndex, row=0)
                    consonantLengthScale.update()
                    canvasWidth += consonantLengthScale.winfo_width() + 2
                    moraHeight += consonantLengthScale.winfo_height()
                    consonantTextLabel: tk.Label = tk.Label(accentPhraseFrame, text=mora.consonant)
                    consonantTextLabel.grid(column=columnIndex, row=1)
                    consonantTextLabel.update()
                    moraHeight += consonantTextLabel.winfo_height()
                    if moraHeight > maxMoraHeight:
                        maxMoraHeight = moraHeight
                    moraHeight = 0
                    columnIndex += 1
                vowelLengthScale: tk.Scale = tk.Scale(accentPhraseFrame, from_=0.3, to=0, resolution=0.01, orient=tk.VERTICAL, command=OnChangeMoraLengthScale(i, j, True))
                vowelLengthScale.set(mora.vowel_length)
                vowelLengthScale.grid(column=columnIndex, row=0)
                vowelLengthScale.update()
                canvasWidth += vowelLengthScale.winfo_width() + 2
                moraHeight += vowelLengthScale.winfo_height()
                vowelTextLabel: tk.Label = tk.Label(accentPhraseFrame, text=mora.vowel)
                vowelTextLabel.grid(column=columnIndex, row=1)
                vowelTextLabel.update()
                moraHeight += vowelTextLabel.winfo_height()
                if moraHeight > maxMoraHeight:
                    maxMoraHeight = moraHeight
                columnIndex += 1
            deleteButton = ttk.Button(accentPhraseFrame, text="削除", command=self.DeleteAccentPhrase(i))
            deleteButton.grid(column=0, row=2, columnspan=columnIndex)
            deleteButton.update()
            maxMoraHeight += deleteButton.winfo_height()
            if maxMoraHeight > canvasHeight:
                canvasHeight = maxMoraHeight
        moraLengthCanvas.configure(scrollregion=(0, 0, canvasWidth, canvasHeight))
        moraLengthCanvas.configure(xscrollcommand=moraLengthScrollbar.set)
        moraLengthCanvas.create_window((0, 0), window=moraLengthInnerFrame, anchor="nw", width=canvasWidth, height=canvasHeight)

    def _DeletePhraseEditor(self) -> None:
        '''
        アクセントフレーズ編集画面の消去
        '''
        if self._intonationFrame is not None:
            for widget in self._intonationFrame.winfo_children():
                widget.destroy()
        if self._moraLengthFrame is not None:
            for widget in self._moraLengthFrame.winfo_children():
                widget.destroy()

    def _MakeVoiceModelList(self) -> None:
        modelListFile = f"{VOICEVOX_PATH}/models/README.txt"
        self.__voiceModelList = {}
        with open(modelListFile, "r", encoding='utf-8') as f:
            for line in f:
                m = re.match(r"\| (\d+.vvm) \| (.+) \| (.+) \| (\d+) \|", line)
                if m:
                    filename: str = m.group(1)
                    charaname: str = m.group(2)
                    modelstyle: str = m.group(3)
                    styleid: str = m.group(4)
                    if not charaname in self.__voiceModelList.keys():
                        self.__voiceModelList[charaname] = {}
                    self.__voiceModelList[charaname][modelstyle] = self.ModelInfo(filename, int(styleid))

    def _LoadVoiceModel(self, character: str, style: str) -> bool:
        if self.__loadedModel:
            for characterMeta in self.__loadedModel.metas:
                if characterMeta.name == character:
                    for styleMeta in characterMeta.styles:
                        if styleMeta.name == style:
                            return False
            # 初期化.
            self.__synthesizer.unload_voice_model(self.__loadedModel.id)
            self.__loadedModel = None
        modelFileDir: str = f"{VOICEVOX_PATH}/models/vvms"
        self.__loadedModel =  VoiceModelFile.open(f"{modelFileDir}/{self.__voiceModelList[character][style].filename}")# type: ignore[attr-defined]
        if self.__loadedModel is None:
            return False
        self.__synthesizer.load_voice_model(self.__loadedModel)
        self.__loadedModel.close()
        return True

    def _UpdateMoraData(self) -> None:
        if self._accentPhrases is not None:
            self._accentPhrases = self.__synthesizer.replace_mora_data(self._accentPhrases, self.__currentStyleID)

    def _UpdateUserDict(self) -> None:
        if self._userDict is not None:
            self._userDict.save(self._user_dict_path)
            self.__open_jtalk.use_user_dict(self._userDict)

    def _UpdateDictionaryEditorList(self, root: tk.Misc | None) -> None:
        if self._listboxFrame is not None:
            listParent: tk.Misc = self._listboxFrame.master
            if root is None:
                root = listParent
            for widget in listParent.winfo_children():
                widget.destroy()
        wordsDic: dict[UUID, voicevox.UserDictWord] | None = None
        if self._userDict is not None:
            wordsDic = self._userDict.to_dict()
        self._listboxFrame = tk.Frame(root)
        self._listboxFrame.pack()
        listbox: tk.Listbox = tk.Listbox(self._listboxFrame, selectmode=tk.SINGLE)
        if wordsDic is not None:
            for word in wordsDic.values():
                listbox.insert(tk.END, f"{word.surface}/ 読み：{word.pronunciation}")
        listbox.insert(tk.END, "追加")
        def OnSelect(_) -> None:
            if listbox is None:
                return
            indices: list[str] = listbox.curselection()
            if len(indices) != 1:
                return
            index = int(indices[0])
            self._accentPhrases = None
            if self._userDict is not None:
                wordsDic = self._userDict.to_dict()
            if wordsDic is None or index >= len(wordsDic):
                self._UpdateDictionaryEditorEdit(None, None)
            else:
                self._UpdateDictionaryEditorEdit(None, list(wordsDic.values())[index])
        listbox.bind("<<ListboxSelect>>", OnSelect)
        listbox.grid(column=0, row=0, sticky=tk.N + tk.S + tk.E + tk.W)
        vbar = tk.Scrollbar(self._listboxFrame, orient=tk.VERTICAL, command=listbox.yview)
        hbar = tk.Scrollbar(self._listboxFrame, orient=tk.HORIZONTAL, command=listbox.xview)
        vbar.grid(column=1, row=0, sticky=tk.N + tk.S)
        hbar.grid(column=0, row=1, sticky=tk.E + tk.W)
        listbox.config(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self._listboxFrame.grid_columnconfigure(0, weight=1)
        self._listboxFrame.grid_rowconfigure(0, weight=1)

        def OnDeletePushed() -> None:
            if self._listboxFrame is None:
                return
            indices: list[str] = listbox.curselection()
            if len(indices) != 1:
                return
            index = int(indices[0])
            if wordsDic is None or index >= len(wordsDic):
                return
            ret =  messagebox.askyesno('', '選択中の単語を削除しますか？')
            if ret:
                userDictWord: voicevox.UserDictWord = list(wordsDic.values())[index]
                self.DelUserDictWord(userDictWord)
        delButton: ttk.Button = ttk.Button(root, text="削除", command=OnDeletePushed)
        delButton.pack()

    def _UpdateDictionaryEditorAccentPhrase(self, root: tk.Misc, accentValue: tk.StringVar) -> None:
        for widget in root.winfo_children():
            widget.destroy()
        
        if self._accentPhrases is None:
            return
        def GetScrollEvent(canvas: tk.Canvas) -> Callable[[tk.Event], None]:
            def inner(event: tk.Event) -> None:
                if event.delta > 0:
                    canvas.xview_scroll(-1, "units")
                elif event.delta < 0:
                    canvas.xview_scroll(1, "units")
            return inner
        accentCanvas: tk.Canvas = tk.Canvas(root)
        accentInnerFrame: tk.Frame = tk.Frame(accentCanvas)
        accentScrollbar: tk.Scrollbar = tk.Scrollbar(root, orient=tk.HORIZONTAL, command=accentCanvas.xview)
        root.grid_columnconfigure(0, weight=1)
        accentCanvas.grid(column=0, row=0, sticky=tk.W + tk.E)
        accentScrollbar.grid(column=0, row=1, sticky=tk.W + tk.E)
        accentCanvas.bind("<MouseWheel>", GetScrollEvent(accentCanvas))
        accentInnerFrame.bind("<MouseWheel>", GetScrollEvent(accentCanvas))
        # 操作部
        def OnChangeAccentButton(accentPhraseIndex: int, value: tk.StringVar) -> Callable[[], None]:
            def inner() -> None:
                if self._accentPhrases is None:
                    return
                self._accentPhrases[accentPhraseIndex].accent = int(value.get())
                self._UpdateMoraData()
                self._UpdatePhraseEditorDisp()
            return inner
        # HACK:canvasのwidth/heightの取り方がよくない。
        canvasWidth: int = 0
        canvasHeight: int = 0
        for i in range(len(self._accentPhrases)):
            accentPhrase: voicevox.AccentPhrase = self._accentPhrases[i]
            columnIndex: int = 0
            accentHeight: int = 0
            for j in range(len(accentPhrase.moras)):
                moraHeight: int = 0
                mora = accentPhrase.moras[j]
                accentRadioButton: ttk.Radiobutton = ttk.Radiobutton(accentInnerFrame, variable=accentValue, value=f"{j+1}", command=OnChangeAccentButton(i, accentValue))
                accentRadioButton.grid(column=columnIndex, row=0)
                accentRadioButton.update()
                canvasWidth += accentRadioButton.winfo_width() + 2
                moraHeight += accentRadioButton.winfo_height()
                textLabel: tk.Label = tk.Label(accentInnerFrame, text=mora.text)
                textLabel.grid(column=columnIndex, row=2)
                textLabel.update()
                moraHeight += textLabel.winfo_height()
                if moraHeight > accentHeight:
                    accentHeight = moraHeight
                columnIndex += 1
            if accentHeight > canvasHeight:
                canvasHeight = accentHeight
        canvasHeight += 2
        accentCanvas.configure(scrollregion=(0, 0, canvasWidth, canvasHeight))
        accentCanvas.configure(xscrollcommand=accentScrollbar.set)
        accentCanvas.create_window((0, 0), window=accentInnerFrame, anchor="nw", width=canvasWidth, height=canvasHeight)

    def _UpdateDictionaryEditorEdit(self, root: tk.Misc | None, userDictWord: voicevox.UserDictWord | None) -> None:
        if self._DictionaryEditFrame is not None:
            if root is None:
                root = self._DictionaryEditFrame.master
            self._DictionaryEditFrame.destroy()
        self._DictionaryEditFrame = tk.Frame(root)
        self._DictionaryEditFrame.pack()
        wordFrame: tk.Frame = tk.Frame(self._DictionaryEditFrame)
        surfaceLabel: ttk.Label = ttk.Label(wordFrame, text="単語")
        surfaceLabel.grid(column=0, row=0)
        surfaceEntry: ttk.Entry = ttk.Entry(wordFrame, width=40)
        if userDictWord is not None:
            surfaceEntry.insert(0, userDictWord.surface)
        surfaceEntry.grid(column=1, row=0)
        pronunciationLabel: ttk.Label = ttk.Label(wordFrame, text="読み(カタカナ)")
        pronunciationLabel.grid(column=0, row=1)
        def Validate(text) -> bool:
            varidatePattern = r"[ァ-ヴー]+"
            result = re.match(varidatePattern, text)
            if result is None or result.start() != 0 or result.end() != len(text):
                return False
            return True
        
        pronunciationEntry: ttk.Entry = ttk.Entry(wordFrame, width=40, validate="key")
        tclValidate = pronunciationEntry.register(Validate)
        pronunciationEntry.config(validatecommand=(tclValidate, "%S"))
        if userDictWord is not None:
            pronunciationEntry.insert(0, userDictWord.pronunciation)
        pronunciationEntry.grid(column=1, row=1)
        wordFrame.pack()
        accentFrame: tk.Frame = tk.Frame(self._DictionaryEditFrame)
        accentValue: tk.StringVar = tk.StringVar(accentFrame)
        if userDictWord is not None:
            accentValue.set(str(userDictWord.accent_type))
        accentPlayButton: ttk.Button = ttk.Button(accentFrame, text="再生")
        accentPlayButton.pack(side=tk.LEFT)
        accentInnerFrame: tk.Frame = tk.Frame(accentFrame)
        accentInnerFrame.pack(side=tk.LEFT)
        self._UpdateDictionaryEditorAccentPhrase(accentInnerFrame, accentValue)
        def OnPlayPushed() -> None:
            charaName: str = list(self.__voiceModelList.keys())[0]
            styleName: str = list(self.__voiceModelList[charaName].keys())[0]
            self.MakeVoice(charaName, styleName, pronunciationEntry.get(), False)
            if self._accentPhrases is None:
                return
            while len(self._accentPhrases) > 1:
                self.MergeAccentPhrase(0)()
            self.MakeVoice(charaName, styleName, pronunciationEntry.get(), False)
            self.PlayWav()
            accentValue.set(str(self._accentPhrases[0].accent))
            self._UpdateDictionaryEditorAccentPhrase(accentInnerFrame, accentValue)
        accentPlayButton["command"] = OnPlayPushed
        accentFrame.pack()

        priorityFrame: tk.Frame = tk.Frame(self._DictionaryEditFrame)
        priorityLabel: ttk.Label = ttk.Label(priorityFrame, text="単語優先度")
        priorityLabel.pack(anchor=tk.W)
        priorityScale: tk.Scale = tk.Scale(priorityFrame, from_=0, to=10, orient=tk.HORIZONTAL)
        if userDictWord is not None:
            priorityScale.set(userDictWord.priority)
        else:
            priorityScale.set(5)
        priorityScale.pack()
        priorityFrame.pack()

        def OnDecide() -> None:
            newUserDictWord = voicevox.UserDictWord(surfaceEntry.get(), pronunciationEntry.get(), int(accentValue.get()), "COMMON_NOUN", int(priorityScale.get()))
            if userDictWord is None:
                self.AddUserDict(newUserDictWord)
            else:
                self.UpdateUserDictWord(self.SearchUserDictWordUUID(userDictWord), newUserDictWord)
            messagebox.showinfo("", "ユーザー辞書を更新しました。")
        decideButton: ttk.Button = ttk.Button(self._DictionaryEditFrame, text="保存", command=OnDecide)
        decideButton.pack(anchor=tk.E)

class PackingData:
    class ElementData:
        def __init__(self, fileName: str) -> None:
            self.__filePath: str = f"{os.environ['RESOLVE_SCRIPT_API']}/{DATA_FILE}/{fileName}"
            self._params: dict[str, Any] = {}
        
        def _InitNewItem(self, paramKey: str, defaultValue: Any) -> None:
            if paramKey not in self._params:
                self._params[paramKey] = defaultValue

        def __getitem__(self, key) -> Any | None:
            if key in self._params:
                return self._params[key]
            return None
        
        def __setitem__(self, key: str, value: Any) -> None:
            if key in self._params and (type(self._params[key]) == type(value) or (type(value) == type(True) and self._params[key] is None)):
                self._params[key] = value
                self._Save()
            else:
                print(f"Invalid key or value type: {key}, {value}, {type(value)} expected {type(self._params.get(key))}")
        
        def Disp(self, frame: tk.Misc, project, trackName: str) -> Any:
            raise NotImplementedError("Subclasses must implement this method")
        
        def DispCheckButton(self, frame: tk.Misc, text: str, key: str) -> tk.Checkbutton:
            '''
            ElementDataに保存したbool値を変更出来るようなラジオボタンを表示する。

            Parameters:
            frame: tk.Misc
                親ウィジェット
            text: ラジオボタンに表示するテキスト
            key: str
                保存したbool値を取り出すキー
            '''
            if key not in self._params:
                print(f"{key}は存在しないため、適当なラジオボタンを作ります。")
                return tk.Checkbutton(frame)
            checkbuttonValue: tk.BooleanVar = tk.BooleanVar()
            def SetFlipX() -> None:
                self[key] = checkbuttonValue.get()
            checkbutton: tk.Checkbutton = tk.Checkbutton(frame, variable=checkbuttonValue, text=text, onvalue=True, offvalue=False, command=SetFlipX)
            if self[key]:
                checkbutton.select()
            else:
                checkbutton.deselect()
            return checkbutton
            

        def _Load(self) -> None:
            os.makedirs(os.path.dirname(os.path.abspath(self.__filePath)), exist_ok=True)
            if os.path.exists(self.__filePath):
                with open(self.__filePath, 'r') as f:
                    data: dict[str, Any] = json.load(f)
                    if data is not None:
                        self._params = data

        def _Save(self) -> None:
            with open(self.__filePath, 'w') as f:
                json.dump(self._params, f, indent=4)
        

    class ImageData(ElementData):
        def __init__(self, fileName: str) -> None:
            super().__init__(fileName)
            self._Load()
            self._InitNewItem("imageDict", {"None": None})
            self._InitNewItem("selectImage", "None")
            self._InitNewItem("x", 0)
            self._InitNewItem("y", 0)
            self._InitNewItem("flipx", False)
            self._InitNewItem("zoom", 1.0)
            self._InitNewItem("voiceOnly", False)
            self._InitNewItem("openedImageDir", "")
            self.imageWidth: int = 200
            self.imageHeight: int = 400
            self.canvas: tk.Canvas | None = None
            self.canvasImage: int | None = None
            self.image: tk.PhotoImage | None = None
            self.combo: ttk.Combobox | None = None

        def __getitem__(self, key) -> Any | None:
            # dictionaryのセーブ回避アクセスを禁止.
            if key == "imageDict":
                return None
            return super().__getitem__(key)

        def getImageDictLen(self) -> int:
            return len(self._params["imageDict"])
        
        def GetImageDictKeys(self) -> list[str]:
            return list(self._params["imageDict"].keys())
        
        def GetImage(self, key: str) -> str:
            return self._params["imageDict"].get(key, "")
        
        def AddImage(self, key: str, path: str) -> None:
            self._params["imageDict"][key] = path
            self._Save()
        
        def DelImage(self, key: str) -> None:
            if key in self._params["imageDict"]:
                del self._params["imageDict"][key]
                self._Save()

        def SetPos(self, x: float | int, y: float | int) -> None:
            self._params["x"] = x
            self._params["y"] = y
            self._Save()
        
        def ApplyToClip(self, clip) -> None:
            '''
            クリップに設定を反映する。

            Parameters:
            clip: timelineClip
                設定を反映させるクリップ
            '''
            # 位置調整
            clip.SetProperty("Pan", self["x"])
            clip.SetProperty("Tilt", self["y"])
            clip.SetProperty("FlipX", self["flipx"])
            clip.SetProperty("ZoomX", self["zoom"])
            clip.SetProperty("ZoomY", self["zoom"])
        
        def Disp(self, frame: tk.Misc, project, trackName: str) -> None:
            '''
            tkinterで画像設定UIを表示する。
            
            Parameters:
            frame: tk.Frame
                UIを表示するフレーム
            project: Resolve,project
                開いているプロジェクト
            trackName: str
                クリップから情報を取得する際の画像トラック名
            '''
            # 画像選択
            self.canvas = tk.Canvas(frame, bg="gray", width=self.imageWidth, height=self.imageHeight)
            self.canvas.pack(fill=tk.BOTH, expand=True)
            self.combo = ttk.Combobox(frame, values=list(self.GetImageDictKeys()))
            self.combo.set(self["selectImage"] if self.GetImageDictKeys() else 'None')
            if self.combo.get() != "None":
                self._ChangeImage(self.GetImage(self.combo.get()))
            def OnSelect(_) -> None:
                if self.combo is None:
                    return
                self["selectImage"] = self.combo.get()
                self._ChangeImage(self.GetImage(self.combo.get()))
            self.combo.bind("<<ComboboxSelected>>", OnSelect) 
            self.combo.pack()
            addImageFrame: ttk.LabelFrame = ttk.LabelFrame(frame, text="表情追加")
            addImageFrame.pack(fill=tk.X, padx=5, pady=5)
            addImageEntry: ttk.Entry = ttk.Entry(addImageFrame)
            addImageEntry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            addImageButton: ttk.Button = ttk.Button(addImageFrame, text="表情追加", command=self._AddImage(addImageEntry))
            addImageButton.pack()
            deleteImageButton: ttk.Button = ttk.Button(frame, text="選択中の表情削除", command=self._DeleteImage)
            deleteImageButton.pack()
            # プロパティ設定
            positionFrame: ttk.LabelFrame = ttk.LabelFrame(frame, text="挿入座標")
            positionFrame.pack(fill=tk.X, padx=5, pady=5)
            xLabel: ttk.Label = ttk.Label(positionFrame, text="X:")
            xLabel.pack(side=tk.LEFT)
            def SetPos() -> bool:
                self.SetPos(float(xEntry.get()), float(yEntry.get()))
                return True
            xEntry: ttk.Entry = ttk.Entry(positionFrame, width=5, validate="focusout", validatecommand=SetPos)
            xEntry.insert(0, str(self["x"]))
            xEntry.pack(side=tk.LEFT, padx=5)
            yLabel: ttk.Label = ttk.Label(positionFrame, text="Y:")
            yLabel.pack(side=tk.LEFT)
            yEntry: ttk.Entry = ttk.Entry(positionFrame, width=5, validate="focusout", validatecommand=SetPos)
            yEntry.insert(0, str(self["y"]))
            yEntry.pack(side=tk.LEFT, padx=5)
            flipXBox: tk.Checkbutton = self.DispCheckButton(positionFrame, "反転", "flipx")
            flipXBox.pack(side=tk.LEFT, padx=5)
            zoomFrame: tk.Frame = tk.Frame(frame)
            zoomFrame.pack(fill=tk.X, padx=5, pady=5)
            zoomLabel: ttk.Label = ttk.Label(zoomFrame, text="拡大率")
            zoomLabel.pack(side=tk.LEFT)
            def SetZoom() -> bool:
                self["zoom"] = float(zoomEntry.get())
                return True
            zoomEntry = ttk.Entry(zoomFrame, width=5, validate="focusout", validatecommand=SetZoom)
            zoomEntry.insert(0, str(self["zoom"]))
            zoomEntry.pack(side=tk.LEFT, padx=5)
            DispVoiceOnlyBox: tk.Checkbutton = self.DispCheckButton(frame, "話している間だけ表示", "voiceOnly")
            DispVoiceOnlyBox.pack(padx=5, pady=5)
            copyButton: ttk.Button = ttk.Button(frame, text="タイムラインから取得", command=self._CopyImageSetting(project, trackName, xEntry, yEntry, flipXBox, zoomEntry))
            copyButton.pack()

        def _ChangeImage(self, imagePath: str) -> None:
            '''
            表示されている画像を変更する
            
            Parameters:
            imagePath: string
                画像のファイルパス
            '''
            if self.canvas is None:
                return
            if imagePath:
                self.image = tk.PhotoImage(file=imagePath)
                imageWidth: int = self.image.width()
                imageHeight: int = self.image.height()
                scale: float = min(self.imageWidth / imageWidth, self.imageHeight / imageHeight, 1)
                if scale < 1:
                    newWidth: int = int(imageWidth * scale)
                    newHeight: int = int(imageHeight * scale)
                    self.image = self.image.subsample(int(imageWidth / newWidth), int(imageHeight / newHeight))
                if self.canvasImage is not None:
                    self.canvas.delete(self.canvasImage)
                self.canvasImage = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.image)
            else:
                if self.canvasImage is not None:
                    self.canvas.delete(self.canvasImage)
                    self.canvasImage = None
            
        def _DeleteImage(self) -> None:
            '''
            選択中の表情を削除する
            '''
            imageName: str | None = self["selectImage"]
            if imageName is None:
                return
            if self.GetImage(imageName):
                self.DelImage(imageName)
                if self.combo is not None:
                    self.combo['values'] = list(self.GetImageDictKeys())
                    if self.combo['values']:
                        self.combo.set(self.combo['values'][0])
                        self._ChangeImage(self.GetImage(self.combo.get()))
                    else:
                        self.combo.set('')
                        if self.canvasImage is not None and self.canvas is not None:
                            self.canvas.delete(self.canvasImage)
                            self.canvasImage = None

        def _AddImage(self, imageNameWidget: tk.Entry) -> Callable[[], None]:
            '''
            表情追加ボタンが押された時の処理を返す。
            表情画像を追加する
            
            Parameters:
            imageNameWidget: tk.Entry
                表情名を入力するウィジェット
            Returns: function
                画像追加ボタンが押されたときに実行される関数
            '''
            def inner() -> None:
                imageName: str = imageNameWidget.get()
                if self.GetImage(imageName):
                    messagebox.showerror("Error", "表情名がすでに存在します。")
                    return
                filePath = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")], initialdir=self["openedImageDir"])
                if filePath:
                    if not imageName :
                        # もう一度表情名を聞く.
                        def OnDestroy() -> None:
                            nonlocal imageName
                            imageName = nameValue.get()
                        askNameWindow: TkinterUtil.SubWindow = TkinterUtil.SubWindow(imageNameWidget, OnDestroy)
                        askNameLabel: ttk.Label = ttk.Label(askNameWindow, text="表情名を入力してください。")
                        askNameLabel.pack()
                        nameValue: tk.StringVar = tk.StringVar()
                        nameValue.set(os.path.splitext(os.path.basename(filePath))[0])
                        askNameEntry: ttk.Entry = ttk.Entry(askNameWindow, textvariable=nameValue)
                        askNameEntry.pack()
                        askNameButtonFrame = tk.Frame(askNameWindow)
                        def OnCancel() -> None:
                            nameValue.set("")
                            askNameWindow.destroy()
                        askNameWindow.protocol('WM_DELETE_WINDOW', OnCancel)
                        askNameCancelButton: ttk.Button = ttk.Button(askNameButtonFrame, text="キャンセル", command=OnCancel)
                        askNameCancelButton.pack(padx=5, pady=5, side=tk.LEFT)
                        def OnDecide() -> None:
                            askNameWindow.destroy()
                        askNameDecideButton: ttk.Button = ttk.Button(askNameButtonFrame, text="決定", command=OnDecide)
                        askNameDecideButton.pack(padx=5, pady=5, side=tk.LEFT)
                        askNameButtonFrame.pack()
                        master = imageNameWidget.master
                        while master.master is not None:
                            master = master.master
                        askNameWindow.grab_set()
                        master.wait_window(askNameWindow)
                        del nameValue
                        if not imageName:
                            # キャンセルされたっぽい
                            return  
                        if self.GetImage(imageName):
                            messagebox.showerror("Error", "表情名がすでに存在します。")
                            return
                    self["openedImageDir"] = os.path.dirname(filePath)
                    self.AddImage(imageName, filePath)
                    if self.combo is None:
                        return
                    self.combo['values'] = list(self.GetImageDictKeys())
                    self.combo.set(imageName)
                    self._ChangeImage(filePath)
            return inner

        def _CopyImageSetting(self, project, trackName: str, xWidget: tk.Entry, yWidget: tk.Entry, flipXWidget: tk.Checkbutton, zoomEntry: tk.Entry) -> Callable[[], None]:
            '''
            現在imageTrackNameトラックで表示されているクリップから情報をコピーする。

            Parameters:
            project: project
                開いているプロジェクト
            trackName: str
                画像トラック名
            xWidget: tkinter.Entry
                x座標の入力ウィジェット
            yWidget: tkinter.Entry
                y座標の入力ウィジェット
            flipXWidget: tkinter.Checkbutton
                x軸反転の入力ウィジェット
            zoomEntry: tkinter.Entry
                拡大率の入力ウィジェット
            '''
            def inner() -> None:
                targetClip = ResolveUtil.GetCurrentTimelineClip(project, TRACK_TYPE_VIDEO_STRING, trackName)
                if not targetClip:
                    messagebox.showerror("ERROR", f"現在表示されている画像が{trackName}に存在しません")
                    return
                ret: bool =  messagebox.askyesno('', '現在表示されているイメージから設定をコピーしますか？')
                if ret:
                    x: int | float = targetClip.GetProperty("Pan")
                    y: int | float = targetClip.GetProperty("Tilt")
                    flipX: bool = targetClip.GetProperty("FlipX")
                    zoom: int | float = targetClip.GetProperty("ZoomX")
                    xWidget.delete(0, tk.END)
                    xWidget.insert(0, str(x))
                    yWidget.delete(0, tk.END)
                    yWidget.insert(0, str(y))
                    zoomEntry.delete(0, tk.END)
                    zoomEntry.insert(0, str(zoom))
                    if flipX and not self["flipX"] or not flipX and self["flipX"]:
                        flipXWidget.invoke()
                    self.SetPos(x, y)
                    self["zoom"] = zoom
            return inner

    class TextData(ElementData):

        def __init__(self, fileName: str, fonts: FontList) -> None:
            super().__init__(fileName)
            self._Load()
            self._InitNewItem("x", 0)
            self._InitNewItem("y", 0)
            self._InitNewItem("font", "")
            self._InitNewItem("style", "")
            self._InitNewItem("size", 0.08)
            self._InitNewItem("color", [1.0, 1.0, 1.0])
            self._InitNewItem("boxWidth", 0.8)
            self._InitNewItem("innerBorderEnabled", False)
            self._InitNewItem("innerBorderThickness", 0.2)
            self._InitNewItem("innerBorderColor", [1.0, 0.0, 0.0])
            self._InitNewItem("outerBorderEnabled", False)
            self._InitNewItem("outerBorderThickness", 0.2)
            self._InitNewItem("outerBorderColor", [1.0, 1.0, 1.0])
            self._InitNewItem("shadowEnabled", False)
            self._InitNewItem("shadowOffset", [0.05, -0.05])
            self._InitNewItem("shadowSize", 1.0)
            self._InitNewItem("shadowColor", [0.0, 0.0, 0.0])

            self.fonts = fonts
            
        def SetPos(self, x: int | float, y: int | float) -> None:
            '''
            位置を設定する。

            Parameters:
            x: int | float
                x座標
            y: int | float
                y座標
            '''
            self._params["x"] = x
            self._params["y"] = y
            self._Save()
        
        def GetSavedColorCode(self, colorType: Literal["color", "innerBorderColor", "outerBorderColor", "shadowColor"]) -> str:
            '''
            保存されている色のカラーコードを取得する。

            Parameters:
            colorType: str
                取得する色のキー
            '''
            colorTuple: tuple[float, float, float] = self._params.get(colorType, (1.0, 1.0, 1.0))
            return GetColorCode(colorTuple[0], colorTuple[1], colorTuple[2])
        
        def ApplyToClip(self, clip) -> None:
            '''
            クリップに設定を反映する。

            Parameters:
            clip: timelineClip
                設定を反映させるクリップ
            '''
            # 位置調整
            clip.SetProperty("Pan", self._params["x"])
            clip.SetProperty("Tilt", self._params["y"])
            
            fusionComp = clip.GetFusionCompByIndex(1)
            textPlus = fusionComp.FindToolByID("TextPlus")
            if textPlus is None:
                messagebox.showerror("ERROR", "TextPlusツールが見つかりませんでした。")
                return
            # フォント設定
            insertFont: str = self.fonts.fonts.get(self._params["font"], "")
            insertStyle: str = self.fonts.style.get(self._params["font"], {}).get(self._params["style"], "")
            textPlus.Font = insertFont
            textPlus.Style = insertStyle
            textPlus.Size = self._params["size"]
            textPlus.Red1 = self._params["color"][0]
            textPlus.Green1 = self._params["color"][1]
            textPlus.Blue1 = self._params["color"][2]
            textPlus.LayoutType = 1
            textPlus.LayoutWidth = self._params["boxWidth"]
            # 内枠設定
            textPlus.Enabled2 = 1 if self._params["innerBorderEnabled"] else 0
            textPlus.Name2 = "InnerOutline"
            textPlus.ElementShape2 = 1
            textPlus.Thickness2 = self._params["innerBorderThickness"]
            textPlus.Red2 = self._params["innerBorderColor"][0]
            textPlus.Green2 = self._params["innerBorderColor"][1]
            textPlus.Blue2 = self._params["innerBorderColor"][2]
            # 外枠設定
            textPlus.Enabled5 = 1 if self._params["outerBorderEnabled"] else 0
            textPlus.Name5 = "OuterOutline"
            textPlus.ElementShape5 = 1
            textPlus.Thickness5 = self._params["innerBorderThickness"] + self._params["outerBorderThickness"]
            textPlus.Red5 = self._params["outerBorderColor"][0]
            textPlus.Green5 = self._params["outerBorderColor"][1]
            textPlus.Blue5 = self._params["outerBorderColor"][2]
            # 影設定
            textPlus.Enabled6 = 1 if self._params["shadowEnabled"] else 0
            textPlus.Name6 = "Shadow"
            textPlus.Offset6 = {1: self._params["shadowOffset"][0], 2: self._params["shadowOffset"][1], 3: 0}
            textPlus.SizeX6 = self._params["shadowSize"]
            textPlus.SizeY6 = self._params["shadowSize"]
            textPlus.Red6 = self._params["shadowColor"][0]
            textPlus.Green6 = self._params["shadowColor"][1]
            textPlus.Blue6 = self._params["shadowColor"][2]
        
        def ColorChooser(self, colorButton: tk.Button, colorType: Literal["color", "innerBorderColor", "outerBorderColor", "shadowColor"]) -> Callable[[], None]:
            '''
            色選択ダイアログを表示し、選択された色を設定する。

            Parameters:
            colorButton: tk.Button
                色を表示するボタンウィジェット
            colorType: string
                設定する色の種類("color", "innerBorderColor", "outerBorderColor", "shadowColor")
            '''
            def inner() -> None:
                color = colorchooser.askcolor((self.GetSavedColorCode(colorType)))
                if color and color[0]:
                    r: float = color[0][0] / 255.0
                    g: float = color[0][1] / 255.0
                    b: float = color[0][2] / 255.0
                    self[colorType] = [r, g, b]
                    hexColor: str = self.GetSavedColorCode(colorType)
                    colorButton["bg"] = hexColor
            return inner

        def Disp(self, frame: tk.Misc, project, trackName: str) -> None:
            '''
            tkinterで字幕設定UIを表示する。
            
            Parameters:
            frame: tk.Frame
                UIを表示するフレーム
            project: Resolve,project
                操作するプロジェクト
            trackName: str
                クリップから情報を取得する際の画像トラック名
            '''
            # 字幕プロパティ
            # フォント情報
            fontFrame: tk.Frame = tk.Frame(frame)
            fontFrame.pack(fill=tk.X, padx=5, pady=5)
            fontLabel: ttk.Label = ttk.Label(fontFrame, text="フォント:")
            fontLabel.pack(side=tk.LEFT)
            fontCombo: ttk.Combobox = ttk.Combobox(fontFrame, values=list(self.fonts.dispFonts))
            fontCombo.set(self["font"] if self["font"] in self.fonts.dispFonts else list(self.fonts.dispFonts)[0])
            def FontChaged(event: tk.Event) -> None:
                self["font"] = fontCombo.get()
                styleCombo['values'] = self.fonts.dispStyles[fontCombo.get()]
                styleCombo.set(self.fonts.dispStyles[fontCombo.get()][0])
            fontCombo.bind("<<ComboboxSelected>>", FontChaged) 
            fontCombo.pack(side=tk.LEFT, padx=5)
            styleLabel: ttk.Label = ttk.Label(fontFrame, text="スタイル:")
            styleLabel.pack(side=tk.LEFT)
            styleCombo: ttk.Combobox = ttk.Combobox(fontFrame, values=self.fonts.dispStyles[fontCombo.get()])
            styleCombo.set(self["style"] if self["style"] in self.fonts.dispStyles[fontCombo.get()] else self.fonts.dispStyles[fontCombo.get()][0])
            def StyleChaged(event: tk.Event) -> None:
                self["style"] = styleCombo.get()
            styleCombo.bind("<<ComboboxSelected>>", StyleChaged) 
            styleCombo.pack(side=tk.LEFT, padx=5)
            # 文字情報
            characterFrame: tk.Frame = tk.Frame(frame)
            characterFrame.pack()
            textPosXLabel: ttk.Label = ttk.Label(characterFrame, text="X:")
            textPosXLabel.pack(side=tk.LEFT)
            def SetPos() -> bool:
                self.SetPos(float(textPosXEntry.get()), float(textPosYEntry.get()))
                return True
            textPosXEntry: ttk.Entry = ttk.Entry(characterFrame, width=5, validate="focusout", validatecommand=SetPos)
            textPosXEntry.insert(0, str(self["x"]))
            textPosXEntry.pack(side=tk.LEFT, padx=5)
            textPosYLabel: ttk.Label = ttk.Label(characterFrame, text="Y:")
            textPosYLabel.pack(side=tk.LEFT)
            textPosYEntry: ttk.Entry = ttk.Entry(characterFrame, width=5, validate="focusout", validatecommand=SetPos)
            textPosYEntry.insert(0, str(self["y"]))
            textPosYEntry.pack(side=tk.LEFT, padx=5)
            sizeFrame: tk.Frame = tk.Frame(characterFrame)
            sizeFrame.pack(side=tk.LEFT, padx=20)
            sizeLabel: ttk.Label = ttk.Label(sizeFrame, text="文字サイズ:")
            sizeLabel.pack(side=tk.LEFT, padx=5)
            def SetSize() -> bool:
                self["size"] = float(sizeEntry.get())
                return True
            sizeEntry: ttk.Entry = ttk.Entry(sizeFrame, width=5, validate="focusout", validatecommand=SetSize)
            sizeEntry.insert(0, str(self["size"]))
            sizeEntry.pack(side=tk.LEFT, padx=5)
            boxWidthLabel: ttk.Label = ttk.Label(sizeFrame, text="文幅")
            boxWidthLabel.pack(side=tk.LEFT, padx=5)
            def SetBoxWidth() -> bool:
                self["boxWidth"] = float(boxWidthEntry.get())
                return True
            boxWidthEntry: ttk.Entry = ttk.Entry(sizeFrame, width=5, validate="focusout", validatecommand=SetBoxWidth)
            boxWidthEntry.insert(0, str(self["boxWidth"]))
            boxWidthEntry.pack(side=tk.LEFT, padx=5)
            colorButton: tk.Button = tk.Button(characterFrame, text="色(RGB)選択")
            colorButton["bg"] = self.GetSavedColorCode("color")
            colorButton.config(command=self.ColorChooser(colorButton, "color"))
            colorButton.pack(side=tk.LEFT, padx=5)
            # 枠線情報
            innerBorderFrame: ttk.LabelFrame = ttk.LabelFrame(frame, text="枠線(内側)設定")
            innerBorderFrame.pack(fill=tk.X, padx=5, pady=5)
            innerBorderEnableCheckBox: tk.Checkbutton = self.DispCheckButton(innerBorderFrame, "有効化:", "innerBorderEnabled")
            innerBorderEnableCheckBox.pack(side=tk.LEFT, padx=5)
            innerBorderSizeLabel: ttk.Label = ttk.Label(innerBorderFrame, text="太さ:")
            innerBorderSizeLabel.pack(side=tk.LEFT)
            def SetInnerBorderSize():
                self["innerBorderThickness"] = float(innerBorderSizeEntry.get())
            innerBorderSizeEntry: ttk.Entry = ttk.Entry(innerBorderFrame, width=5, validate="focusout", validatecommand=SetInnerBorderSize)
            innerBorderSizeEntry.insert(0, str(self["innerBorderThickness"]))
            innerBorderSizeEntry.pack(side=tk.LEFT, padx=5)
            innerBorderColorButton: tk.Button = tk.Button(innerBorderFrame, text="色(RGB)選択")
            innerBorderColorButton["bg"] = self.GetSavedColorCode("innerBorderColor")
            innerBorderColorButton.config(command=self.ColorChooser(innerBorderColorButton, "innerBorderColor"))
            innerBorderColorButton.pack(side=tk.LEFT, padx=5)
            outerBorderFrame: ttk.LabelFrame = ttk.LabelFrame(frame, text="枠線(外側)設定")
            outerBorderFrame.pack(fill=tk.X, padx=5, pady=5)
            outerBorderEnableCheckBox: tk.Checkbutton = self.DispCheckButton(outerBorderFrame, "有効化:", "outerBorderEnabled")
            outerBorderEnableCheckBox.pack(side=tk.LEFT, padx=5)
            outerBorderSizeLabel: ttk.Label = ttk.Label(outerBorderFrame, text="太さ:")
            outerBorderSizeLabel.pack(side=tk.LEFT)
            def SetOuterBorderSize():
                self["outerBorderThickness"] = float(outerBorderSizeEntry.get())
            outerBorderSizeEntry: ttk.Entry = ttk.Entry(outerBorderFrame, width=5, validate="focusout", validatecommand=SetOuterBorderSize)
            outerBorderSizeEntry.insert(0, str(self["outerBorderThickness"]))
            outerBorderSizeEntry.pack(side=tk.LEFT, padx=5)
            outerBorderColorButton: tk.Button = tk.Button(outerBorderFrame, text="色(RGB)選択")
            outerBorderColorButton["bg"] = self.GetSavedColorCode("outerBorderColor")
            outerBorderColorButton.config(command=self.ColorChooser(outerBorderColorButton, "outerBorderColor"))
            outerBorderColorButton.pack(side=tk.LEFT, padx=5)
            # 影情報
            shadowFrame: ttk.LabelFrame = ttk.LabelFrame(frame, text="影設定")
            shadowFrame.pack(fill=tk.X, padx=5, pady=5)
            shadowEnableCheckBox: tk.Checkbutton = self.DispCheckButton(shadowFrame, "有効化:", "shadowEnabled")
            shadowEnableCheckBox.pack(side=tk.LEFT, padx=5)
            shadowOffsetFrame: ttk.LabelFrame = ttk.LabelFrame(shadowFrame, text="位置")
            shadowOffsetFrame.pack(side=tk.LEFT, padx=5)
            shadowOffsetXLabel: ttk.Label = ttk.Label(shadowOffsetFrame, text="x")
            shadowOffsetXLabel.pack(side=tk.LEFT, padx=5)
            def SetShadowOffset() -> bool:
                self["shadowOffset"] = [float(shadowOffsetXEntry.get()), float(shadowOffsetYEntry.get())]
                return True
            shadowOffsetXEntry: ttk.Entry = ttk.Entry(shadowOffsetFrame, width=5, validate="focusout", validatecommand=SetShadowOffset)
            shadowOffsetXEntry.insert(0, str(cast(list[float], self["shadowOffset"])[0]))
            shadowOffsetXEntry.pack(side=tk.LEFT, padx=5)
            shadowOffsetYLabel: ttk.Label = ttk.Label(shadowOffsetFrame, text="y")
            shadowOffsetYLabel.pack(side=tk.LEFT, padx=5)
            shadowOffsetYEntry: ttk.Entry = ttk.Entry(shadowOffsetFrame, width=5, validate="focusout", validatecommand=SetShadowOffset)
            shadowOffsetYEntry.insert(0, str(cast(list[float], self["shadowOffset"])[1]))
            shadowOffsetYEntry.pack(side=tk.LEFT, padx=5)
            shadowSizeLabel: ttk.Label = ttk.Label(shadowFrame, text="影文字サイズ")
            shadowSizeLabel.pack(side=tk.LEFT, padx=5)
            def SetShadowSize() -> bool:
                self["shadowSize"] = float(shadowSizeEntry.get())
                return True
            shadowSizeEntry: ttk.Entry = ttk.Entry(shadowFrame, width=5, validate="focusout", validatecommand=SetShadowSize)
            shadowSizeEntry.insert(0, str(self["shadowSize"]))
            shadowSizeEntry.pack(side=tk.LEFT, padx=5)
            shadowColorButton: tk.Button = tk.Button(shadowFrame, text="影色選択")
            shadowColorButton["bg"] = self.GetSavedColorCode("shadowColor")
            shadowColorButton.config(command=self.ColorChooser(shadowColorButton, "shadowColor"))
            shadowColorButton.pack(side=tk.LEFT, padx=5)

            textCopyButton: ttk.Button = ttk.Button(frame, text="タイムラインから取得", command=self._CopyTextSetting(project, trackName, textPosXEntry, textPosYEntry, fontCombo, styleCombo, sizeEntry, boxWidthEntry, colorButton, 
                                                                                                            innerBorderEnableCheckBox,  innerBorderSizeEntry, innerBorderColorButton, outerBorderEnableCheckBox,  outerBorderSizeEntry, outerBorderColorButton,
                                                                                                            shadowEnableCheckBox, shadowOffsetXEntry, shadowOffsetYEntry, shadowSizeEntry, shadowColorButton))
            textCopyButton.pack()

        def _CopyTextSetting(self, project, trackName: str, xWidget: tk.Entry, yWidget: tk.Entry, fontCombo: ttk.Combobox, styleCombo: ttk.Combobox, sizeEntry: tk.Entry, boxWidthEntry: tk.Entry, colorButton: tk.Button, 
                             innerBorderEnabledWidget: tk.Checkbutton, innerBorderThicknessEntry: tk.Entry, innerBorderColorButton: tk.Button, outerBorderEnabledWidget: tk.Checkbutton, outerBorderThicknessEntry: tk.Entry, outerBorderColorButton: tk.Button, 
                             shadowEnabledWidget: tk.Checkbutton, shadowOffsetXEntry: tk.Entry, shadowOffsetYEntry: tk.Entry, shadowSizeEntry: tk.Entry, shadowColorButton: tk.Button) -> Callable[[], None]:
            '''
            現在textTrackNameトラックで表示されているクリップから情報をコピーする。

            Parameters:
            project: Resolve,project
                操作するプロジェクト
            trackName: str
                字幕トラック名
            xWidget: tkinter.Entry
                x座標の入力ウィジェット
            yWidget: tkinter.Entry
                y座標の入力ウィジェット
            fontCombo: ttk.Combobox
                フォント名の入力ウィジェット
            styleCombo: ttk.Combobox
                フォントスタイルの入力ウィジェット
            sizeEntry: tkinter.Entry
                フォントサイズの入力ウィジェット
            colorButton: tkinter.Button
                文字色の入力ウィジェット
            innerBorderEnabledWidget: tkinter.Checkbutton
                内枠有効化の入力ウィジェット
            innerBorderThicknessEntry: tkinter.Entry
                内枠太さの入力ウィジェット
            innerBorderColorButton: tkinter.Button
                内枠色の入力ウィジェット
            outerBorderEnabledWidget: tkinter.Checkbutton
                外枠有効化の入力ウィジェット
            outerBorderThicknessEntry: tkinter.Entry
                外枠太さの入力ウィジェット
            outerBorderColorButton: tkinter.Button
                外枠色の入力ウィジェット
            '''
            def inner() -> None:
                targetClip = ResolveUtil.GetCurrentTimelineClip(project, TRACK_TYPE_VIDEO_STRING, trackName)
                if not targetClip:
                    messagebox.showerror("ERROR", f"現在表示されている字幕が{trackName}に存在しません")
                    return
                ret: bool =  messagebox.askyesno('', '現在表示されている字幕から設定をコピーしますか？')
                if ret:
                    x: int | float = targetClip.GetProperty("Pan")
                    y: int | float = targetClip.GetProperty("Tilt")
                    fusionComp = targetClip.GetFusionCompByIndex(1)
                    textPlus = fusionComp.FindToolByID("TextPlus")
                    print(textPlus.GetInput("LayoutWidth")) 
                    if textPlus is None:
                        messagebox.showerror("ERROR", "TextPlusツールが見つかりませんでした。")
                        return
                    dispFont: str = ""
                    dispStyle: str = ""
                    font: str = textPlus.GetInput("Font")
                    style: str = textPlus.GetInput("Style")
                    for fontName, fontObj in self.fonts.fonts.items():
                        if fontObj == font:
                            dispFont = fontName
                            break
                    for styleName, styleObj in self.fonts.style.get(dispFont, {}).items():
                        if styleObj == style:
                            dispStyle = styleName
                            break
                    size: int | float = textPlus.GetInput("Size")
                    color: list[float] = [textPlus.GetInput("Red1"), textPlus.GetInput("Green1"), textPlus.GetInput("Blue1")]
                    if textPlus.GetInput("LayoutType") == 1:
                        boxWidth = textPlus.GetInput("LayoutWidth")
                    else:
                        boxWidth = 0.8
                    innerBorderEnabled: int = textPlus.GetInput("Enabled2")
                    innerBorderThickness: int | float = textPlus.GetInput("Thickness2")
                    innerBorderColor: list[float] = [textPlus.GetInput("Red2"), textPlus.GetInput("Green2"), textPlus.GetInput("Blue2")]
                    outerBorderEnabled: int = textPlus.GetInput("Enabled5")
                    if innerBorderEnabled:
                        outerBorderThickness: int | float = textPlus.GetInput("Thickness5") - innerBorderThickness
                    else:
                        outerBorderThickness = textPlus.GetInput("Thickness5")
                    outerBorderColor: list[float] = [textPlus.GetInput("Red5"), textPlus.GetInput("Green5"), textPlus.GetInput("Blue5")]
                    # 影設定
                    shadowEnabled: int = textPlus.GetInput("Enabled6")
                    shadowOffset: dict[int, float] = textPlus.GetInput("Offset6")
                    shadowSize: float = textPlus.GetInput("SizeX6")
                    shadowColor: list[float] = [textPlus.GetInput("Red6"), textPlus.GetInput("Green6"), textPlus.GetInput("Blue6")]
                    xWidget.delete(0, tk.END)
                    xWidget.insert(0, str(x))
                    yWidget.delete(0, tk.END)
                    yWidget.insert(0, str(y))
                    self.SetPos(x, y)
                    fontCombo.set(dispFont)
                    self["font"] = dispFont
                    styleCombo.set(dispStyle)
                    self["style"] = dispStyle
                    sizeEntry.delete(0, tk.END)
                    sizeEntry.insert(0, str(size))
                    self["size"] = size
                    boxWidthEntry.delete(0, tk.END)
                    boxWidthEntry.insert(0, str(boxWidth))
                    self["boxWidth"] = boxWidth
                    self["color"] = color
                    colorButton["bg"] = self.GetSavedColorCode("color")
                    self["boxWidth"] = boxWidth
                    if innerBorderEnabled and not self["innerBorderEnabled"] or not innerBorderEnabled and self["innerBorderEnabled"]:
                        innerBorderEnabledWidget.invoke()
                    innerBorderThicknessEntry.delete(0, tk.END)
                    innerBorderThicknessEntry.insert(0, str(innerBorderThickness))
                    self["innerBorderThickness"] = innerBorderThickness
                    self["innerBorderColor"] = innerBorderColor
                    innerBorderColorButton["bg"] = self.GetSavedColorCode("innerBorderColor")
                    if outerBorderEnabled and not self["outerBorderEnabled"] or not outerBorderEnabled and self["outerBorderEnabled"]:
                        outerBorderEnabledWidget.invoke()
                    outerBorderThicknessEntry.delete(0, tk.END)
                    outerBorderThicknessEntry.insert(0, str(outerBorderThickness))
                    self["outerBorderThickness"] = outerBorderThickness
                    self["outerBorderColor"] = outerBorderColor
                    outerBorderColorButton["bg"] = self.GetSavedColorCode("outerBorderColor")
                    # 影設定
                    if shadowEnabled and not self["shadowEnabled"] or not shadowEnabled and self["shadowEnabled"]:
                        shadowEnabledWidget.invoke()
                    shadowOffsetXEntry.delete(0, tk.END)
                    shadowOffsetXEntry.insert(0, str(shadowOffset[1]))
                    shadowOffsetYEntry.delete(0, tk.END)
                    shadowOffsetYEntry.insert(0, str(shadowOffset[2]))
                    self["shadowOffset"] = [shadowOffset[1], shadowOffset[2]]
                    shadowSizeEntry.delete(0, tk.END)
                    shadowSizeEntry.insert(0, str(shadowSize))
                    self["shadowSize"] = shadowSize
                    self._params["shadowColor"] = shadowColor
                    shadowColorButton["bg"] = self.GetSavedColorCode("shadowColor")
            return inner

    class VoicevoxData(ElementData):
        def __init__(self, fileName: str, voicevox: VoicevoxEngine) -> None:
            super().__init__(fileName)
            self._Load()
            self.voicevox: VoicevoxEngine = voicevox
            self._InitNewItem("character", voicevox.GetCharacterList()[0])
            self._InitNewItem("style", voicevox.GetStyleList(self._params["character"])[0])
            self._InitNewItem("upspeak", False)
            self._InitNewItem("outDir", "")
            self._InitNewItem("speed", 1.0)
            self._InitNewItem("pitch", 0.0)
            self._InitNewItem("intonation", 1.0)
            self._InitNewItem("volume", 1.0)
            self._InitNewItem("pauseLengthScale", 1.0)
            self._InitNewItem("prePhonemeLength", 0.1)
            self._InitNewItem("postPhonemeLength", 0.1)
        
        def Disp(self, frame: tk.Misc, project, trackName: str) -> None:
            '''
            tkinterでvoicevox設定UIを表示する。
            
            Parameters:
            frame: tk.Frame
                UIを表示するフレーム
            project: Resolve,project
                使わない引数
            trackName: str
                使わない引数
            '''
            voiceOption: tk.Frame = tk.Frame(frame, width=400, height=50)
            voiceOption.pack()
            character: ttk.Combobox = ttk.Combobox(voiceOption, values=self.voicevox.GetCharacterList())
            character.set(self["character"] if self["character"] in self.voicevox.GetCharacterList() else self.voicevox.GetCharacterList()[0])
            character.pack(side=tk.LEFT)
            style: ttk.Combobox = ttk.Combobox(voiceOption, values=self.voicevox.GetStyleList(character.get()))
            style.set(self["style"] if self["style"] in self.voicevox.GetStyleList(character.get()) else self.voicevox.GetStyleList(character.get())[0])
            style.pack(side=tk.LEFT, padx=5)
            def OnCharacterSelected(_: tk.Event) -> None:
                self["character"] = character.get()
                style['values'] = self.voicevox.GetStyleList(character.get())
                style.set(self.voicevox.GetStyleList(character.get())[0])
            character.bind("<<ComboboxSelected>>", OnCharacterSelected)
            def OnStyleSelected(_: tk.Event) -> None:
                self["style"] = style.get()
            style.bind("<<ComboboxSelected>>", OnStyleSelected)
            upspeakBox = self.DispCheckButton(voiceOption, "疑問文時の文末上げ:", "upspeak")
            upspeakBox.pack(side=tk.LEFT, padx=5)
            voicePropertyFrame: tk.Frame = tk.Frame(frame)
            voicePropertyFrame.pack()
            def VoicePropertyDisp(frame: tk.Misc, text: str, key: str, from_: float, to_: float, column: int, row: int) -> None:
                propertyLabel: tk.Label = tk.Label(frame, text=text)
                propertyLabel.grid(column=column, row=row)
                def OnScaleClicked(_: str) -> None:
                    self[key] = propertyScale.get()
                propertyScale: tk.Scale = tk.Scale(frame, from_=from_, to=to_, orient=tk.HORIZONTAL, command=OnScaleClicked, resolution=0.01)
                propertyScale.set(self[key])
                propertyScale.grid(column=column+1, row=row)
            VoicePropertyDisp(voicePropertyFrame, "話速", "speed", 0.5, 2.0, 0, 0)
            VoicePropertyDisp(voicePropertyFrame, "音高", "pitch", -0.15, 0.15, 2, 0)
            VoicePropertyDisp(voicePropertyFrame, "抑揚", "intonation", 0.0, 2.0, 4, 0)
            VoicePropertyDisp(voicePropertyFrame, "音量", "volume", 0.0, 2.0, 0, 1)
            VoicePropertyDisp(voicePropertyFrame, "間の長さ", "pauseLengthScale", 0.0, 2.0, 2, 1)
            VoicePropertyDisp(voicePropertyFrame, "開始無音", "prePhonemeLength", 0.0, 1.5, 0, 2)
            VoicePropertyDisp(voicePropertyFrame, "終了無音", "postPhonemeLength", 0.0, 1.5, 2, 2)
            VoicevoxDirectorySelectFrame: tk.Frame = tk.Frame(frame)
            VoicevoxDirectorySelectFrame.pack(anchor = tk.W)
            voicevoxDirectorySelectButton: ttk.Button = ttk.Button(VoicevoxDirectorySelectFrame, text="出力先フォルダ選択", command=self.SelectVoicevoxOutFolder)
            voicevoxDirectorySelectButton.pack(side=tk.LEFT)
            self.voicevoxDirectoryLabel: ttk.Label = ttk.Label(VoicevoxDirectorySelectFrame)
            self.voicevoxDirectoryLabel["text"] = self["outDir"]
            self.voicevoxDirectoryLabel.pack(side=tk.LEFT, padx=5)

        def SelectVoicevoxOutFolder(self) -> None:
            '''
            voicevoxの音声ファイル出力先フォルダを選択する
            '''
            dirPath: str = filedialog.askdirectory(initialdir=self["outDir"])
            if dirPath:
                self["outDir"] = dirPath
                self.voicevoxDirectoryLabel["text"] = dirPath
        
    def __init__(self, name: str, project, fonts: FontList) -> None:
        '''
        Parameters:
        name: string
            パック名
        project: project
            操作するプロジェクト
        fonts: FontList
            インストール済みフォントリスト
        '''
        global voicevoxAvailable
        self.name: str = name
        self.project = project
        self.voiceTrackName: str = f"{self.name}Voice"
        self.imageTrackName: str = f"{self.name}Image"
        self.textTrackName: str = f"{self.name}Text"
        self.imageData: PackingData.ImageData = self.ImageData(f"{name}_image.json")
        self.textData: PackingData.TextData = self.TextData(f"{name}_text.json", fonts)
        self.trackLockStatus: dict[tuple[str, int], bool] = {}
        self.openedVoiceDir: str = ""
        if voicevoxAvailable:
            self.voicevox: VoicevoxEngine = VoicevoxEngine()
            if not self.voicevox.IsInitSucceeded():
                messagebox.showerror("ERROR", "voicevoxの初期化に失敗しました")
                voicevoxAvailable = False
            self.voicevoxData = self.VoicevoxData(f"{name}_voicevox.json", self.voicevox)
        pass
    
    def SelectTrack(self, timeline, trackType: Literal["video", "audio"], trackName: str, exec: Callable[[int], Any] | None=None) -> bool:
        '''
        指定された名前のトラックを選択し、callbackの処理を実行する。存在しない場合は新規にトラックを作成する
        ロックされていない、最もインデックスの小さいトラックが選択されるようなので、それを利用する

        Parameters:
        timeline: timeline
            操作するタイムライン
        trackType: "video" or "audio"
            トラックの種類
        trackName: string
            トラック名
        exec: function
            トラック選択後に実行する関数

        Returns: bool
            成功したかどうか
        '''
        if not timeline:
            messagebox.showerror("Error", "有効なタイムラインがありません。")
            return False
        trackIndex: int = -1
        for type in TRACK_TYPES:
            if not type:
                continue
            trackCount: int = timeline.GetTrackCount(type)
            for i in range(1, trackCount + 1):
                if timeline.GetTrackName(type, i) == trackName and type == trackType:
                    if(timeline.GetIsTrackLocked(type, i)):
                        messagebox.showerror("Error", f"{trackType}トラック '{trackName}' はロックされています。")
                        self.RevertTrackLock(timeline)
                        return False
                    trackIndex = i
                else:
                    self.trackLockStatus[(type, i)] = timeline.GetIsTrackLocked(type, i)
                    timeline.SetTrackLock(type, i, True)
        if trackIndex == -1:
            AddTrackResult: bool = timeline.AddTrack(trackType)
            trackIndex = timeline.GetTrackCount(trackType)
            if not AddTrackResult:
                messagebox.showerror("Error", f"{trackType}トラックの追加に失敗しました。")
                return False
            trackCount = timeline.GetTrackCount(trackType)
            timeline.SetTrackName(trackType, trackCount, trackName)
        if exec:
            exec(trackIndex)
        self.RevertTrackLock(timeline)
        return True
    
    def RevertTrackLock(self, timeline) -> None:
        '''
        SelectTrackで変更したトラックのロック状態を元に戻す

        Parameters:
        timeline: timeline
            操作するタイムライン
        trackType: "video" or "audio"
            トラックの種類
        '''
        for trackdata, lockStatus in self.trackLockStatus.items():
            timeline.SetTrackLock(trackdata[0], trackdata[1], lockStatus)
        self.trackLockStatus = {}

    def GetTemplateClipFromMediaPool(self, mediaPoolPath: str) -> Any | None:
        '''
        メディアプールから空のFusionCompテンプレートクリップを取得する
        存在しない場合は新規に作成する

        Returns: MediaPoolItem
            取得したテンプレートクリップ
        '''
        if not self.project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return None
        clipName: str = f"{CLIP_NAME_PREFIX}Template"
        # メディアプールフォルダの指定
        mediaPool = self.project.GetMediaPool()
        prevCurrentFolder = mediaPool.GetCurrentFolder()
        ResolveUtil.MoveCurrentFolder(self.project.GetMediaPool(), mediaPoolPath)
        currentFolder = mediaPool.GetCurrentFolder()
        # 既存のテンプレートクリップを探す
        clips = currentFolder.GetClips()
        retClip = None
        for clip in clips.values():
            if clip.GetName() == clipName:
                retClip = clip
                break
        if not retClip:
            # テンプレートクリップが存在しない場合は新規に作成する
            currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(self.project)
            def exec(_):
                newClip = currentTimeline.InsertFusionCompositionIntoTimeline()
                if not newClip:
                    messagebox.showerror("Error", "テンプレートクリップの作成に失敗しました。")
                    return
                fusionClip = currentTimeline.CreateFusionClip(newClip)
                currentTimeline.DeleteClips([fusionClip])
            if self.SelectTrack(currentTimeline, TRACK_TYPE_VIDEO_STRING, self.textTrackName, exec):
                retClip = currentFolder.GetClips()[len(currentFolder.GetClips())]
                retClip.SetName(clipName)
        mediaPool.SetCurrentFolder(prevCurrentFolder)
        return retClip

    def GetClipFromMediaPoolWithFilePath(self, filePath: str, clipName: str, mediaPoolPath) -> Any | None:
        '''
        メディアプールから指定されたファイルパスのクリップを取得する
        存在しない場合は新規に作成する

        Parameters:
        filePath: string
            クリップのファイルパス
        clipName: string
            クリップ名

        Returns: MediaPoolItem
            取得したクリップ
        '''
        if not self.project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return None
        mediaPool = self.project.GetMediaPool()
        prevCurrentFolder = mediaPool.GetCurrentFolder()
        ResolveUtil.MoveCurrentFolder(self.project.GetMediaPool(), mediaPoolPath)
        currentFolder = mediaPool.GetCurrentFolder()
        clips = currentFolder.GetClips()
        for clip in clips.values():
            if clip.GetName() == clipName:
                mediaPool.SetCurrentFolder(prevCurrentFolder)
                return clip
        # クリップが存在しない場合は新規に作成する
        newClip = mediaPool.ImportMedia([filePath])
        if not newClip or len(newClip) == 0:
            messagebox.showerror("Error", f"クリップ '{filePath}' のインポートに失敗しました。")
            mediaPool.SetCurrentFolder(prevCurrentFolder)
            return None
        newClip = newClip[0]
        newClip.SetName(clipName)
        mediaPool.SetCurrentFolder(prevCurrentFolder)
        return newClip
    
    def InsertVoice(self, waveFile: str) -> None:
        '''
        現在のタイムラインに音声を挿入する
        タイムラインが存在しない場合は新規に作成する

        Parameters:
        waveFile: str
            挿入する音声waveファイルのパス
        '''
        if not waveFile:
            messagebox.showerror("Error", "有効な音声ファイルがありません")
            return
        if not self.project:
            messagebox.showerror("Error", "有効なプロジェクトがありません")
            return
        currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(self.project)

        def exec(trackIndex: int) -> None:
            clip = self.GetClipFromMediaPoolWithFilePath(waveFile, os.path.splitext(os.path.basename(waveFile))[0], f"/{CLIP_NAME_PREFIX}/Voices")
            fps: int | float = currentTimeline.GetSetting("timelineFrameRate")
            currentTime: str = currentTimeline.GetCurrentTimecode()
            mediaPool = self.project.GetMediaPool()
            offset: int = ResolveUtil.TimecodeToFrames(currentTime, fps)
            newClips = mediaPool.AppendToTimeline([{
                "mediaPoolItem": clip,
                "startFrame": 0,
                "trackIndex": trackIndex,
                "mediaType": TRACK_TYPE_AUDIO,
                "recordFrame": offset
            }])
            if newClips is None or len(newClips) == 0 or not newClips[0]:
                messagebox.showerror("Error", "音声の挿入に失敗しました。")
                return
        
        self.SelectTrack(currentTimeline, TRACK_TYPE_AUDIO_STRING, self.voiceTrackName, exec)
    
    def InsertFusionClip(self, timeline, endtimecode: str, trackIndex: int, mediaType: int) -> Any | None:
        '''
        指定したタイムラインの現在タイムコードに、FusionClipを配置する。

        Parameters:
        timeline: Timeline
            指定するタイムライン
        endtimecode: str(hh:mm:dd:ff)
            終了タイムコード
        trackIndex: int
            配置するトラック
        mediaType: int
            配置するトラックのメディアタイプ
        '''
        fps: int | float = timeline.GetSetting("timelineFrameRate")
        currentTime = timeline.GetCurrentTimecode()
        mediaPool = self.project.GetMediaPool()
        offset: int = ResolveUtil.TimecodeToFrames(currentTime, fps)
        fusionClip = self.GetTemplateClipFromMediaPool(f"/{CLIP_NAME_PREFIX}/Templates")
        if fusionClip is None:
            return None
        clipFps: int = fusionClip.GetClipProperty("FPS")
        duration: float = ResolveUtil.TimecodeToFrames(endtimecode, clipFps) - ResolveUtil.TimecodeToFrames(currentTime, clipFps)
        newClips = mediaPool.AppendToTimeline([{
            "mediaPoolItem": fusionClip,
            "startFrame": 0,
            "endFrame": duration,
            "trackIndex": trackIndex,
            "mediaType": mediaType,
            "recordFrame": offset
        }])
        if newClips is None or len(newClips) == 0 or not newClips[0]:
            return None
        return newClips[0]
    
    def ReinsertImage(self, clip, starttimecode: str, endtimecode: str) -> None:
        '''
        画像クリップを再挿入する
        開始タイムコード/終了タイムコードを再設定するために使用する。

        Parameters:
        clip: timelineClip
            再挿入するクリップ
        starttimecode: str
            クリップの開始タイムコード
        endtimecode: str
            クリップの終了タイムコード
        '''
        if not self.project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return
        currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(self.project)

        def exec(trackIndex: int) -> None:
            # 既存クリップの情報取得
            oldName: str = clip.GetName()
            trackType, trackIndex = clip.GetTrackTypeAndIndex()
            oldLoaderTool = clip.GetFusionCompByIndex(1).FindToolByID("Loader")
            oldfile = oldLoaderTool.GetInput("Clip")
            oldtrim = oldLoaderTool.GetInput("ClipTimeStart")
            oldx = clip.GetProperty("Pan")
            oldy = clip.GetProperty("Tilt")
            oldflipx = clip.GetProperty("FlipX")
            oldzoom = clip.GetProperty("ZoomX")
            currentTimeline.DeleteClips([clip])
            # 画像クリップの再挿入
            currentTimeline.SetCurrentTimecode(starttimecode)
            newImage = self.InsertFusionClip(currentTimeline, endtimecode, trackIndex, trackType)
            if newImage is None:
                messagebox.showerror("Error", "画像の挿入に失敗しました。")
                return 
            newImage.SetName(oldName)
            # 画像の設定
            if newImage.GetFusionCompCount() == 0:
                newImage.AddFusionComp()
            fusionComp = newImage.GetFusionCompByIndex(1)
            fusionComp.Lock()
            loaderTool = fusionComp.AddTool("Loader", 0, 0)
            loaderTool.Clip = oldfile
            loaderTool.ClipTimeStart = oldtrim
            loaderTool.ClipTimeEnd = oldtrim
            loaderTool.Loop = 1.0
            fusionComp.Unlock()
            # 表示
            mediaOut = fusionComp.FindToolByID("MediaOut")
            if not mediaOut:
                mediaOut = fusionComp.AddTool("MediaOut", 1000, 0)
            # プロパティ反映
            newImage.SetProperty("Pan", oldx)
            newImage.SetProperty("Tilt", oldy)
            newImage.SetProperty("FlipX", oldflipx)
            newImage.SetProperty("ZoomX", oldzoom)
            newImage.SetProperty("ZoomY", oldzoom)
            mediaOut.Input = loaderTool.Output

        self.SelectTrack(currentTimeline, TRACK_TYPE_VIDEO_STRING, self.imageTrackName, exec)

    def InsertImage(self, endtimecode: str) -> None:
        '''
        現在のタイムラインに画像を挿入する
        タイムラインが存在しない場合は新規に作成する
        画像が選択されていない場合は何もしない
        
        Parameters:
        endtimecode: str
            画像クリップの終了タイムコード
        '''
        file: str = self.imageData.GetImage(cast(str, self.imageData["selectImage"]))        
        if not file:
            return
        if not self.project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return
        currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(self.project)

        currentClip = ResolveUtil.GetCurrentTimelineClip(self.project, TRACK_TYPE_VIDEO_STRING, self.imageTrackName)
        if currentClip is not None:
            # 挿入位置に既存クリップが存在する場合、現在時間までのクリップとして置き直す.
            oldClipStartFrame = currentClip.GetStart(False)
            fps: int | float = currentTimeline.GetSetting("timelineFrameRate")
            starttimecode: str = ResolveUtil.GetTimecodeFromFrame(oldClipStartFrame, fps)
            self.ReinsertImage(currentClip, starttimecode, currentTimeline.GetCurrentTimecode())
            
        def exec(trackIndex: int) -> None:
            newImage = self.InsertFusionClip(currentTimeline, endtimecode, trackIndex, TRACK_TYPE_VIDEO)
            if newImage is None:
                messagebox.showerror("Error", "画像の挿入に失敗しました。")
                return 
            newImage.SetName(f"{self.name}Image_{self.imageData['selectImage']}")
            # 画像の設定
            if newImage.GetFusionCompCount() == 0:
                newImage.AddFusionComp()
            fusionComp = newImage.GetFusionCompByIndex(1)
            fusionComp.Lock()
            loaderTool = fusionComp.AddTool("Loader", 0, 0)
            loaderTool.Clip = file
            fusionComp.Unlock()
            # 表示画像の固定
            # ファイル名末尾が数字だと、自動的に1つのアニメーションにされてしまうので、そのアニメーションの1コマを指定してそこで固定させる
            trim = 0
            m = re.match(r"(.*)(\d+)\.(\w+)", file)
            if m:
                trim = int(m.group(2))
                for i in range(0, trim):
                    if len(glob.glob(f"{m.group(1)}*{i}.{m.group(3)}")) == 0:
                        trim -= 1
                        break
            loaderTool.ClipTimeStart = trim
            loaderTool.ClipTimeEnd = trim
            loaderTool.Loop = 1.0
            # 表示
            mediaOut = fusionComp.FindToolByID("MediaOut")
            if not mediaOut:
                mediaOut = fusionComp.AddTool("MediaOut", 1000, 0)
            # プロパティ反映
            self.imageData.ApplyToClip(newImage)
            mediaOut.Input = loaderTool.Output

        self.SelectTrack(currentTimeline, TRACK_TYPE_VIDEO_STRING, self.imageTrackName, exec)

    def InsertText(self, text: str, endtimecode: str) -> None:
        '''
        現在のタイムラインにテキストを挿入する
        タイムラインが存在しない場合は新規に作成する
        
        Parameters:
        text: string
            挿入するテキスト
        endtimecode: string
            テキストクリップの終了タイムコード
        '''
        if not text:
            messagebox.showerror("Error", "テキストが空です。")
            return
        if not self.project:
            messagebox.showerror("Error", "有効なプロジェクトがありません。")
            return
        currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(self.project)

        def exec(trackIndex):
            newtext = self.InsertFusionClip(currentTimeline, endtimecode, trackIndex, TRACK_TYPE_VIDEO)
            if newtext is None:
                messagebox.showerror("Error", "字幕の挿入に失敗しました。")
                return 
            newtext.SetName(f"{self.name}Text_{text[:10]}")
            if newtext.GetFusionCompCount() == 0:
                newtext.AddFusionComp()
            fusionComp = newtext.GetFusionCompByIndex(1)
            # 文字列の挿入
            textTool = fusionComp.AddTool("TextPlus", 0, 0)
            textTool.StyledText = text
            mediaOut = fusionComp.FindToolByID("MediaOut")
            if not mediaOut:
                mediaOut = fusionComp.AddTool("MediaOut", 1000, 0)
            mediaOut.Input = textTool.Output
            # プロパティ反映
            self.textData.ApplyToClip(newtext)

        self.SelectTrack(currentTimeline, TRACK_TYPE_VIDEO_STRING, self.textTrackName, exec)
        
    def InsertVoicevox(self, textWidget: tk.Text) -> Callable[[], None]:
        '''
        Voicevoxモードで挿入ボタンが押された時の処理を返す。

        Parameters:
        textWidget: tk.Text
            テキストウィジェット

        Returns: function
            挿入ボタンが押されたときに実行される関数
        '''
        def inner() -> None:
            text: str = textWidget.get("1.0", tk.END)
            if text == "\n":
                messagebox.showerror("Error", "テキストが空です。")
                return
            filename: str = text.replace('\n', '')
            filepath: str = f"{self.voicevoxData['outDir']}/{filename[:min(10, len(filename))]}.wav"
            fixedFilepath: str = filepath
            counter: int = 0
            while os.path.exists(fixedFilepath):
                counter += 1
                fixedFilepath = filepath[:-4] + str(counter) + ".wav"
            self.voicevox.MakeVoice(cast(str, self.voicevoxData["character"]), cast(str, self.voicevoxData["style"]), text, cast(bool, self.voicevoxData["upspeak"]), cast(float, self.voicevoxData["speed"]), cast(float, self.voicevoxData["pitch"]), cast(float, self.voicevoxData["intonation"]), cast(float, self.voicevoxData["volume"]), cast(float, self.voicevoxData["pauseLengthScale"]), cast(float, self.voicevoxData["prePhonemeLength"]), cast(float, self.voicevoxData["postPhonemeLength"]))
            self.voiceDuration["text"] = f"{self.voicevox.CalcWavDuration():.2f}秒"
            self.voicevox.SaveWav(fixedFilepath)
            self.InsertRaw(fixedFilepath, text)
        return inner
    
    def InsertExistFile(self) -> None:
        '''
        既存ファイルのデータからテキスト・画像・音声をタイムライン上に挿入する
        '''
        voiceFile: str = self.voiceFileLabel["text"]
        textFile: str = voiceFile[:voiceFile.rfind('.')] + ".txt"
        text: str = ""
        if not self.textEnableValue.get():
            text = ""
        elif not os.path.exists(textFile):
            ret = messagebox.askyesno("", f"字幕データが見つかりませんでした。字幕なしで挿入しますか？\n検索したファイルパス：{textFile}")
            if not ret:
                messagebox.showinfo("", "音声挿入を中断しました。")
                return
        else:
            with open(textFile, encoding="utf-8") as f:
                text = "".join(f.readlines())
        self.InsertRaw(voiceFile, text)

    def InsertRaw(self, wavFile, text: str) -> None:
        '''
        テキスト・画像・音声を現在のタイムラインに挿入する

        Parameters:
        wavFile: str
            挿入する音声データのパス
        text: str
            字幕として挿入するテキスト
        '''
        currentTimeline = ResolveUtil.GetOrCreateCurrentTimeline(self.project)
        currentTimecode = currentTimeline.GetCurrentTimecode()
        self.InsertVoice(wavFile)
        endTimecode: str = currentTimeline.GetCurrentTimecode()
        currentTimeline.SetCurrentTimecode(currentTimecode)
        # デフォルトでは最終フレームまで画像を表示する
        imageEndTimecode: str = ResolveUtil.GetTimecodeFromFrame(currentTimeline.GetEndFrame(), currentTimeline.GetSetting("timelineFrameRate"))
        if self.imageData["voiceOnly"]:
            imageEndTimecode = endTimecode
        self.InsertImage(imageEndTimecode)
        if self.textEnableValue.get() and text and len(text) > 0:
            currentTimeline.SetCurrentTimecode(currentTimecode)
            self.InsertText(text, endTimecode)
        currentTimeline.SetCurrentTimecode(endTimecode)
    
    def PlayVoicevox(self, textWidget: tk.Text) -> Callable[[], None]:
        '''
        Voicevoxの音声を鳴らしてみる。

        Parameters:
        textWidget: tk.Text
            再生する文章が書いてあるウィジェット
        
        Returns: function
            再生ボタンが押された時に実行される関数
        '''
        def inner() -> None:
            text: str = textWidget.get("1.0", tk.END)
            if text == "\n":
                messagebox.showerror("Error", "テキストが空です。")
                return
            self.voicevox.MakeVoice(cast(str, self.voicevoxData["character"]), cast(str, self.voicevoxData["style"]), text, cast(bool, self.voicevoxData["upspeak"]), cast(float, self.voicevoxData["speed"]), cast(float, self.voicevoxData["pitch"]), cast(float, self.voicevoxData["intonation"]), cast(float, self.voicevoxData["volume"]), cast(float, self.voicevoxData["pauseLengthScale"]), cast(float, self.voicevoxData["prePhonemeLength"]), cast(float, self.voicevoxData["postPhonemeLength"]))
            self.voicevox.PlayWav()
            self.voiceDuration["text"] = f"{self.voicevox.CalcWavDuration():.2f}秒"
        return inner

    def SelectExistVoice(self) -> None:
        filePath: str = filedialog.askopenfilename(filetypes=[("音声ファイル", "*.wav")], initialdir=self.openedVoiceDir)
        if filePath:
            self.openedVoiceDir = os.path.dirname(filePath)
            self.voiceFileLabel["text"] = filePath

    def Disp(self, root: tk.Misc) -> None:
        '''
        ウィジェットを表示する

        Parameters:
        root: tk.Tk or tk.Frame
            ウィジェットを配置する親ウィジェット
        '''
        panedWindow: ttk.PanedWindow = ttk.Panedwindow(root, orient=tk.HORIZONTAL)

        # 左側
        leftFrame: tk.Frame = tk.Frame(panedWindow, width=200, height=300, relief=tk.SUNKEN)
        self.imageData.Disp(leftFrame, self.project, self.imageTrackName)
        panedWindow.add(leftFrame, weight=1) 

        # 右側
        rightFrame: tk.Frame = tk.Frame(panedWindow, width=400, height=300, relief=tk.SUNKEN)
        self.textData.Disp(rightFrame, self.project, self.textTrackName)

        # 字幕無効化チェックボックス
        self.textEnableValue: tk.BooleanVar = tk.BooleanVar()
        textEnableCheckBox: tk.Checkbutton = tk.Checkbutton(rightFrame, variable=self.textEnableValue, text="字幕を挿入する", onvalue=True, offvalue=False)
        textEnableCheckBox.select()
        textEnableCheckBox.pack()

        # 音声
        voiceNote: ttk.Notebook = ttk.Notebook(rightFrame)
        if voicevoxAvailable:
            # voicevox出力タブ
            voicevoxFrame: tk.Frame = tk.Frame(voiceNote)
            voiceNote.add(voicevoxFrame, text="voicevox出力")
            self.voicevoxData.Disp(voicevoxFrame, None, "")
            text: tk.Text = tk.Text(voicevoxFrame, height=3, width=50)
            text.pack()
            testbutton: ttk.Button = ttk.Button(voicevoxFrame, text="再生", command=self.PlayVoicevox(text))
            testbutton.pack()
            self.voicevox.InitPhraseEditorDisp(voicevoxFrame)
            insertFrame: tk.Frame = tk.Frame(voicevoxFrame, width=400, height=20)
            insertFrame.pack(fill=tk.X)
            voicevoxInsertButton: ttk.Button = ttk.Button(insertFrame, text="挿入", command=self.InsertVoicevox(text))
            voicevoxInsertButton.pack(side=tk.RIGHT)
            self.voiceDuration: ttk.Label = ttk.Label(insertFrame, text="0秒")
            self.voiceDuration.pack(side=tk.RIGHT)
        # 既存ファイル使用タブ
        fileFrame: tk.Frame = tk.Frame(voiceNote)
        voiceNote.add(fileFrame, text="既存ファイル使用")
        voiceNote.pack(fill=tk.BOTH, expand=True, padx=5)
        fileSelectFrame: tk.Frame = tk.Frame(fileFrame)
        fileSelectFrame.pack(anchor = tk.W)
        voiceFileButton: ttk.Button = ttk.Button(fileSelectFrame, text="ファイル選択", command=self.SelectExistVoice)
        voiceFileButton.pack(side=tk.LEFT)
        self.voiceFileLabel: ttk.Label = ttk.Label(fileSelectFrame, text="")
        self.voiceFileLabel.pack(side=tk.LEFT, padx=5)
        fileInsertButton: ttk.Button = ttk.Button(fileFrame, text="挿入", command=self.InsertExistFile)
        fileInsertButton.pack()

        panedWindow.add(rightFrame, weight=3)
        panedWindow.pack(fill=tk.BOTH, expand=True)

def AddTemplateInFile(name, filePath: str) -> None:
    '''
    テンプレート名をファイルに追加する。
    テンプレート名がすでに存在する場合はエラー終了

    Parameters:
    name: str
        テンプレート名
    filePath: str
        保存先
    '''
    with open(filePath, "a+", encoding="utf-8") as f:
        f.seek(0, 0)
        for templateName in f:
            templateName = templateName.replace("\n", "")
            if templateName == name:
                messagebox.showerror("ERROR", "その名前はすでに存在します。")
                return
        f.seek(0, 2)
        f.write(name + "\n")

def OpenAddTemplateGUI(templateFile: str, root: tk.Tk |  tk.Toplevel | None = None, OnDestroy: Callable[[], None] | None = None):
    '''
    キャラ追加UIを開く。

    Parameters:
    templateFile: str
        追加キャラを書き込むファイル名
    root: tkinter.Widget
        ルート。NoneでなければToplebelで作成
    OnDestroy: Function() => None
        UIが閉じたあとの処理を追加。

    Returns: Boolean
        キャラを追加したらTrue
    '''
    global templateRoot
    if root is None:
        templateRoot = tk.Tk()
    else:
        if templateRoot is not None and templateRoot.winfo_exists():
            return False
        templateRoot = TkinterUtil.SubWindow(root, OnDestroy)
        templateRoot.transient(root)
    templateRoot.title("キャラ追加")
    label: ttk.Label = ttk.Label(templateRoot, text="キャラ名を入力してください")
    label.pack()
    entry: ttk.Entry = ttk.Entry(templateRoot)
    def OnButtonPushed():
        AddTemplateInFile(entry.get(), templateFile)
        templateRoot.destroy()
    entry.pack()
    button: ttk.Button = ttk.Button(templateRoot, text="作成", command=OnButtonPushed)
    button.pack()
    if root is None:
        templateRoot.mainloop()
    return True

def AddTemplate(root: tk.Tk | tk.Toplevel, templateFile: str, notebook: ttk.Notebook, project, fonts: FontList) -> Callable[[], None]:
    '''
    キャラを追加する関数を返す。

    Parameters:
    root: tkinter.Widget
        ルート
    templateFile: str
        追加キャラを書き込むファイル名
    notebook: ttk.notebook
        追加後にタブを追加するnotebook
    project: 
        PackingDataで扱うProject
    fonts: FontList
        インストールされているフォント情報

    Returns: function
        キャラ追加をする関数
    '''
    def inner() -> None:
        def OnDestroy() -> None:
            if os.path.exists(templateFile):
                with open(templateFile, "r", encoding="utf-8") as f:
                    templates: list[str] = f.readlines()
                if len(templates) > len(notebook.tabs()):
                    AddTab(notebook, templates[-1].replace("\n", ""), project, fonts)
        OpenAddTemplateGUI(templateFile, root, OnDestroy)
    return inner

def AddTab(notebook: ttk.Notebook, templateName: str, project, fonts: FontList) -> None:
    '''
    notebookにPackingData情報のタブを追加する

    Parameters:
    notebook: ttk.Notebook
        タブを追加するnotebook
    templateName: str
        PackingDataの名前
    project: Resolve.Project
        PackingDataで扱うProject
    '''
    tab: tk.Frame = tk.Frame(notebook)
    notebook.add(tab, text=templateName)
    displayData: PackingData = PackingData(templateName, project, fonts)
    displayData.Disp(tab)

def GetGithubReleasesLatestName(owner: str, repo: str) -> tuple[str, str]:
    try:
        try:
            with urllib.request.urlopen(f"https://api.github.com/repos/{owner}/{repo}/releases/latest", timeout=5) as response:
                data = json.loads(response.read().decode('utf-8'))
                retName: str = data.get('name', '')
                retTagName: str = data.get('tag_name', '')
                if not re.match(r"^\d+\.\d+\.\d$", retName):
                    retName = ''
        except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError, KeyError, TimeoutError):
            retName = ''
            retTagName = ''
    except:
        retName = ''
        retTagName = ''
    return retName, retTagName

def CompareVersion(version: str, target: str) -> bool:
    if version == '' or target == '':
        return False
    majorV, minorV, patchV = re.findall(r"\d+", version)
    majorT, minorT, patchT = re.findall(r"\d+", target)
    if int(majorV) > int(majorT):
        return True
    elif int(majorV) < int(majorT):
        return False
    elif int(minorV) > int(minorT):
        return True
    elif int(minorV) < int(minorT):
        return False
    elif int(patchV) > int(patchT):
        return True
    return False

def VersionCheck() -> bool:
    global scriptVersion
    if os.path.exists(IGNORE_VERSION_FILE):
        with open(IGNORE_VERSION_FILE, "r") as f:
            ignoreVersion = f.readline()
            if CompareVersion(ignoreVersion, scriptVersion):
                scriptVersion = ignoreVersion
    LATEST: Final = GetGithubReleasesLatestName("GlintAugly", "VoiceInserter")
    ret: bool = True
    TEMPFILE: Final = f"{os.environ['RESOLVE_SCRIPT_API']}/{DATA_FILE}/temp.bat"
    if os.path.exists(TEMPFILE):
        os.remove(TEMPFILE)
    if voicevoxAvailable:
        if CompareVersion(VOICEVOX_TARGET_VERSION, voicevox.__version__):
            askYesNo: bool = messagebox.askyesno("", "VOICEVOX_COREのバージョンが要求バージョンと異なります。アップデートしますか？")
            if askYesNo:
                with open(TEMPFILE, "w", encoding="shift_jis") as f:
                    f.writelines(['@echo off\n',
                                'setlocal enabledelayedexpansion\n',
                                f'set PID={os.getpid()}\n',
                                'timeout /t 2 /nobreak\n',
                                ':wait_loop\n',
                                'tasklist /fi "PID eq %PID%" | find /i "%PID%" >nul\n',
                                'if %ERRORLEVEL% == 1 (\n'
                                '    echo プロセスID %PID% は終了しました\n',
                                '    goto end\n',
                                ')\n',
                                'echo プロセスID %PID% はまだ実行中です...\n',
                                'timeout /t 5 /nobreak\n',
                                'goto wait_loop\n',
                                ':end\n',
                                f'pip install https://github.com/VOICEVOX/voicevox_core/releases/download/{VOICEVOX_TARGET_VERSION}/voicevox_core-{VOICEVOX_TARGET_VERSION}-cp310-abi3-win_amd64.whl --prefix="%RESOLVE_SCRIPT_API%"\\Modules\\voicevox_core\n',
                                'echo MsgBox "アップデートが完了しました。VoiceInserterを再度起動してください。",vbInformation,"info" > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs\n',
                                'del /Q %TEMP%\msgbox.vbs\n',
                                'exit 0'])
                try:
                    subprocess.run(f'start "" "{TEMPFILE}"', shell=True)
                except subprocess.CalledProcessError as e:
                    messagebox.showerror("エラー", "アップデートに失敗しました。")
                    print(f"Error: Command '{e.cmd}' returned non-zero exit status {e.returncode}")
                    print(f"Output: {e.stdout}")
                    print(f"Error Output: {e.stderr}")
                return False

    if CompareVersion(LATEST[0], scriptVersion):
        updateroot: tk.Tk = tk.Tk()
        updateroot.title("アップデート確認")
        updateLabel: tk.Label = tk.Label(updateroot, text=f"VoiceInserterのバージョン更新を検知しました。\n{LATEST[0]}にアップデートしますか？")
        updateLabel.pack()
        updateButtonFrame: tk.Frame = tk.Frame(updateroot)
        updateYesButton: ttk.Button = ttk.Button(updateButtonFrame, text="はい")
        def OnPressYes() -> None:
            with open(TEMPFILE, "w", encoding="shift_jis") as f:
                f.writelines(['@echo off\n',
                            'setlocal enabledelayedexpansion\n',
                            f'set PID={os.getpid()}\n',
                            'timeout /t 2 /nobreak\n',
                            ':wait_loop\n',
                            'tasklist /fi "PID eq %PID%" | find /i "%PID%" >nul\n',
                            'if %ERRORLEVEL% == 1 (\n',
                            '    echo プロセスID %PID% は終了しました\n',
                            '    goto end\n',
                            ')\n',
                            'echo プロセスID %PID% はまだ実行中です...\n',
                            'timeout /t 5 /nobreak\n',
                            'goto wait_loop\n',
                            ':end\n',
                            f'curl -L -o "{os.path.dirname(sys.argv[0])}\\temp.txt" https://github.com/GlintAugly/VoiceInserter/releases/download/{LATEST[1]}/VoiceInserter.py\n',
                            f'findstr /R "^Python Script$" "{os.path.dirname(sys.argv[0])}\\temp.txt"\n'
                            'if %ERRORLEVEL% == 0 (\n',
                            f'    move /y "{os.path.dirname(sys.argv[0])}\\temp.txt" "{sys.argv[0]}"\n',
                            ') else (\n',
                            '    echo MsgBox "アップデートファイルのダウンロードに失敗しました。",vbCritical,"error" > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs\n'
                            f'    del /Q "{os.path.dirname(sys.argv[0])}\\temp.txt"\n',
                            '    del /Q %TEMP%\msgbox.vbs\n',
                            '    exit 1\n',
                            ')\n',
                            'echo MsgBox "アップデートが完了しました。VoiceInserterを再度起動してください。",vbInformation,"info" > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs\n',
                            'del /Q %TEMP%\msgbox.vbs\n',
                            'exit 0'])
            try:
                subprocess.run(f'start "" "{TEMPFILE}"', shell=True)
            except subprocess.CalledProcessError as e:
                messagebox.showerror("エラー", "アップデートに失敗しました。")
                print(f"Error: Command '{e.cmd}' returned non-zero exit status {e.returncode}")
                print(f"Output: {e.stdout}")
                print(f"Error Output: {e.stderr}")

            nonlocal ret
            ret = False
            updateroot.destroy()
        updateYesButton.config(command=OnPressYes)
        updateYesButton.pack(side=tk.LEFT)
        updateNoButton: ttk.Button = ttk.Button(updateButtonFrame, text="いいえ")
        def OnPressNo() -> None:
            updateroot.destroy()
        updateNoButton.config(command=OnPressNo)
        updateNoButton.pack(side=tk.LEFT)
        updateIgnoreButton: ttk.Button = ttk.Button(updateButtonFrame, text="このバージョンを無視する")
        def OnPressIgnore() -> None:
            with open(IGNORE_VERSION_FILE, "w") as f:
                f.write(LATEST[0])
            updateroot.destroy()
        updateIgnoreButton.config(command=OnPressIgnore)
        updateIgnoreButton.pack(side=tk.LEFT)
        updateButtonFrame.pack()
        updateroot.mainloop()
    return ret

if __name__ == "__main__":
    # バージョンチェック
    if not VersionCheck():
        sys.exit()
    # projectの取得
    resolve = app.GetResolve() # type: ignore
    projectManager = resolve.GetProjectManager()
    project = projectManager.GetCurrentProject()
    installedFonts: FontList = FontList(FontList.FetchFonts())
    
    templateFile: str = f"{os.environ['RESOLVE_SCRIPT_API']}/{DATA_FILE}/templates.dat"
    os.makedirs(os.path.dirname(os.path.abspath(templateFile)), exist_ok=True)
    if not os.path.exists(templateFile):
        # 初期設定
        OpenAddTemplateGUI(templateFile)
        if not os.path.exists(templateFile):
            sys.exit()
            
    # GUI表示
    root: tk.Tk = tk.Tk()
    root.title("Voice Inserter")
    # タブ追加
    notebook: ttk.Notebook = ttk.Notebook(root)
    with open(templateFile, "r", encoding="utf-8") as f:
        for templateName in f:
            templateName = templateName.replace("\n", "")
            AddTab(notebook, templateName, project, installedFonts)
    notebook.pack(fill='both', expand=True)
    # メニューバー追加
    menuBar: tk.Menu = tk.Menu(root, tearoff=0)
    fileMenu: tk.Menu = tk.Menu(menuBar, tearoff=0)
    templateRoot: tk.Tk | tk.Toplevel | None = None
    fileMenu.add_command(label="キャラ追加", command=AddTemplate(root, templateFile, notebook, project, installedFonts))
    menuBar.add_cascade(label="file", menu=fileMenu)
    root.config(menu=menuBar)
    root.mainloop()
