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
    /// 
    /// </summary>
    [SerializeField]
    List<DMXDataStruct> Data = new List<DMXDataStruct>();

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
                    //Debug.Log(Table);

                    DataRowCollection DataRow = Result.Tables[Table.ToString()].Rows;
                    //Debug.LogWarning("Rows : " + DataRow.Count);

                    DMXDataStruct Data = new DMXDataStruct(short.Parse(DataRow[0].ItemArray[1].ToString()));

                    #region 記下資料要到哪裡中止讀取
                    int DmxEnd = 0;
                    for (int a = 3; a < DataRow[0].ItemArray.Length; a++)
                    {
                        if (int.TryParse(DataRow[0].ItemArray[a].ToString(), out int DMXID))
                        {
                            Debug.Log("DMX ID : " + DMXID);
                        }
                        else
                        {
                            DmxEnd = a;
                            break;
                        }
                    }
                    #endregion

                    //foreach (DataRow Row in DataRow)
                    //{
                    //    //Debug.Log(string.Join(", ", Row.ItemArray));

                    //    //設Universe
                    //    if (Row.ItemArray[0].ToString() == "Universe")
                    //    {
                    //        //Debug.Log("Universe : " + Row.ItemArray[1]);

                    //        Data = new DMXDataStruct(short.Parse(Row.ItemArray[1].ToString()));

                    //        for (int a = 3; a < Row.ItemArray.Length; a++)
                    //        {
                    //            if (int.TryParse(Row.ItemArray[a].ToString(), out int DMXID))
                    //            {
                    //                Debug.Log("DMX ID : " + DMXID);
                    //            }
                    //            else
                    //            {
                    //                DmxEnd = a;
                    //                break;
                    //            }
                    //        }
                    //    }



                    //    //if (TimeSpan.TryParseExact(Row.ItemArray[1].ToString(), new string[] { @"mm\:ss", @"m\:ss" }, null, out TimeSpan ActionTime))
                    //    //{
                    //    //    Data.Add(new DataStruct(ActionTime, new float[] {
                    //    //        float.Parse(Row.ItemArray[2].ToString()) ,
                    //    //        float.Parse(Row.ItemArray[3].ToString()) ,
                    //    //        float.Parse(Row.ItemArray[4].ToString()) ,
                    //    //        float.Parse(Row.ItemArray[5].ToString()),
                    //    //        float.Parse(Row.ItemArray[6].ToString()),
                    //    //        float.Parse(Row.ItemArray[7].ToString())
                    //    //    }));

                    //    //    //Debug.Log(ID + " ActionTime : " + ActionTime.TotalSeconds + " (" + ActionTime.Minutes.ToString("00") + ":" + ActionTime.Seconds.ToString("00") + ")");

                    //    //    //ID++;
                    //    //}
                    //}

                    #region 讀取時間與數值
                    for (int a = 2; a < DmxEnd; a++)//每一行
                    {
                        Debug.LogWarning(string.Join(", ", DataRow[a].ItemArray));

                        for (int b = 0; b < DmxEnd; b++)
                        {
                            Debug.Log(a + "-" + b + " : " + DataRow[a].ItemArray[b]);
                        }
                    }
                    #endregion


                }

                //var ID = 0;



            }
        }

    }

    // Update is called once per frame
    void Update()
    {
        if (AVProPlayer.Control.IsPlaying())
        {
            var CurrentTime = AVProPlayer.Control.GetCurrentTime();

            var TS = TimeSpan.FromSeconds(CurrentTime);
            Txt_Time1.text = TS.Minutes.ToString("00") + ":" + TS.Seconds.ToString("00");

            Txt_Time2.text = ((int)CurrentTime).ToString();
        }
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
    /// </summary>
    public DMXValueStruct[] Values;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="Universe"></param>
    public DMXDataStruct(short Universe)
    {
        this.Universe = Universe;
        Values = new DMXValueStruct[512];
    }
}
/// <summary>
/// 
/// </summary>
[Serializable]
struct DMXValueStruct
{
    /// <summary>
    /// 
    /// </summary>
    public TimeSpan ActionTime;
    /// <summary>
    /// 
    /// </summary>
    public float Value;
    /// <summary>
    /// 是否已觸發
    /// </summary>
    public bool IsTrigger;

    public DMXValueStruct(TimeSpan ActionTime, float Value)
    {
        this.ActionTime = ActionTime;
        this.Value = Value;
        IsTrigger = false;
    }
}