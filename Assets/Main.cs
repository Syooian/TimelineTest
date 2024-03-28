using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;
using ExcelDataReader;
using System.Data;
using System;
using RenderHeads.Media.AVProVideo;
//using System.Globalization;
//using static System.Data.IDataReader;

public class Main : MonoBehaviour
{
    /// <summary>
    /// DMX資料
    /// <para>Key : Universe</para>
    /// </summary>
    [SerializeField]
    Dictionary<short, DMXDataStruct> DMXData = new Dictionary<short, DMXDataStruct>();

    /// <summary>
    /// 
    /// </summary>
    [SerializeField]
    MediaPlayer AVProPlayer;

    /// <summary>
    /// 影片時間顯示
    /// </summary>
    [SerializeField]
    Text Txt_Time1;
    /// <summary>
    /// 影片時間顯示 (總秒數)
    /// </summary>
    [SerializeField]
    Text Txt_Time2;

    // Start is called before the first frame update
    void Start()
    {
        using (var Stream = File.Open(Application.streamingAssetsPath + "/Test2.xlsx", FileMode.Open, FileAccess.Read))
        {
            using (var Reader = ExcelReaderFactory.CreateReader(Stream))
            {
                var Result = Reader.AsDataSet();

                Debug.LogWarning("資料表數量 : " + Result.Tables.Count);

                foreach (var Table in Result.Tables)
                {
                    Debug.Log("資料表名稱：" + Table);

                    DataRowCollection DataRow = Result.Tables[Table.ToString()].Rows;
                    Debug.LogWarning("Rows : " + DataRow.Count);

                    short Universe = short.Parse(DataRow[0].ItemArray[1].ToString());
                    DMXDataStruct UniverseData;
                    if (DMXData.ContainsKey(Universe))//如果資料內已有該Universe的資料，先取出繼續增加
                        UniverseData = DMXData[Universe];
                    else
                        UniverseData = new DMXDataStruct(Universe);

                    Debug.LogWarning("Universe : " + Universe);

                    #region 記下資料要到哪裡中止讀取
                    //如果Excel檔有整理乾淨，後面沒有不必要的值，其實不需要這段，直接取ItemArray的長度即可

                    int DmxEnd = 0;
                    Debug.Log("Row0 Length : " + DataRow[0].ItemArray.Length);
                    Debug.Log(string.Join(", ", DataRow[0].ItemArray));
                    for (int a = 3; a < DataRow[0].ItemArray.Length; a++)
                    {
                        if (int.TryParse(DataRow[0].ItemArray[a].ToString(), out int DMXID))
                        {
                            Debug.Log("DMX ID : " + DMXID);

                            DmxEnd = a;
                        }
                        else
                        {

                            break;
                        }
                    }

                    Debug.Log("DmxEnd : " + DmxEnd);
                    #endregion

                    #region 讀取時間與數值
                    for (int a = 2; a < DataRow.Count; a++)//每一行
                    {
                        Debug.LogWarning(string.Join(", ", DataRow[a].ItemArray));

                        //總秒數
                        var ActionTime = (int)TimeSpan.FromSeconds(int.Parse(DataRow[a].ItemArray[0].ToString())).TotalSeconds;
                        //Debug.Log(ActionTime.TotalSeconds + " - " + ToDateTimeString(ActionTime));

                        //找出現有的時間軸資料，沒有就直接New一個
                        ActionTimeStruct Timeline;
                        if (UniverseData.ActionTime.ContainsKey(ActionTime))
                            Timeline = UniverseData.ActionTime[ActionTime];
                        else
                        {
                            Timeline = new ActionTimeStruct();
                            Timeline.DevicesValue = new byte[512];
                        }

                        for (int b = 3; b <= DmxEnd; b++)
                        {
                            Debug.Log(a + "-" + b + " : " + DataRow[a].ItemArray[b]);
                            Timeline.DevicesValue[int.Parse(DataRow[0].ItemArray[b].ToString())] = byte.Parse(DataRow[a].ItemArray[b].ToString());
                        }

                        UniverseData.ActionTime[ActionTime] = Timeline;

                        //Debug.Log(ActionTime.TotalSeconds + " - " + ToDateTimeString(ActionTime) + " = " + string.Join(", ", ats.DevicesValue));
                    }
                    #endregion

                    DMXData[Universe] = UniverseData;
                }
            }
        }

        //foreach (var Data in DMXData)
        //{
        //    Debug.LogWarning("Universe : " + Data.Value.Universe);

        //    foreach (var Timeline in Data.Value.ActionTime)
        //    {
        //        Debug.Log(
        //            Timeline.Key + " = " +
        //            string.Join(", ", Timeline.Value.DevicesValue));
        //    }
        //}
    }

    /// <summary>
    /// 紀錄上次的時間
    /// <para>總秒數</para>
    /// </summary>
    int LastCurrentTime = -1;

    // Update is called once per frame
    void Update()
    {
        if (AVProPlayer.Control.IsPlaying())
        {
            var CurrentTime = AVProPlayer.Control.GetCurrentTime();

            Txt_Time1.text = ToDateTimeString(TimeSpan.FromSeconds(CurrentTime));

            var NewCurrentTime = (int)CurrentTime;

            Txt_Time2.text = NewCurrentTime.ToString();

            if (LastCurrentTime != NewCurrentTime)
            {
                LastCurrentTime = NewCurrentTime;

                //以當前播放時間查詢該時間是否需觸發DMX執行
                foreach (var ActionData in DMXData)
                {
                    if (ActionData.Value.ActionTime.ContainsKey(LastCurrentTime))//有紀錄該觸發時間
                    {
                        Debug.Log(Txt_Time2.text + " - " + Txt_Time1.text + " = " + string.Join(", ", ActionData.Value.ActionTime[LastCurrentTime].DevicesValue));
                    }
                }
            }
        }
    }

    string ToDateTimeString(TimeSpan TS)
    {
        return TS.Minutes.ToString("00") + ":" + TS.Seconds.ToString("00");
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="Player"></param>
    /// <param name="EventType"></param>
    /// <param name="ErrorCode"></param>
    public void AVProEvents(MediaPlayer Player, MediaPlayerEvent.EventType EventType, ErrorCode ErrorCode)
    {

    }
}

/// <summary>
/// 
/// </summary>
[Serializable]
struct DMXDataStruct
{
    /// <summary>
    /// 
    /// </summary>
    public short Universe;
    /// <summary>
    /// 
    /// <para>Key : 觸發時間總秒數</para>
    /// </summary>
    public Dictionary<int, ActionTimeStruct> ActionTime;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="Universe"></param>
    public DMXDataStruct(short Universe)
    {
        this.Universe = Universe;
        ActionTime = new Dictionary<int, ActionTimeStruct>();
    }
}
/// <summary>
/// 
/// </summary>
[Serializable]
struct ActionTimeStruct
{
    /// <summary>
    /// 
    /// </summary>
    public byte[] DevicesValue;
    /// <summary>
    /// 是否已觸發
    /// </summary>
    public bool IsTrigger;
}