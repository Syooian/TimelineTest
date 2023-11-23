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
    List<DataStruct> Data = new List<DataStruct>();

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
    IEnumerator Start()
    {
        using (var Stream = File.Open(Application.streamingAssetsPath + "/Test2.xlsx", FileMode.Open, FileAccess.Read))
        {
            using (var Reader = ExcelReaderFactory.CreateReader(Stream))
            {
                var Result = Reader.AsDataSet();

                //var ID = 0;

                DataRowCollection DataRow = Result.Tables["工作表1"].Rows;
                //Debug.LogWarning("Rows : " + DataRow.Count);

                foreach (DataRow Row in DataRow)
                {
                    //Debug.LogWarning(string.Join(", ", Row.ItemArray));
                    if (TimeSpan.TryParseExact(Row.ItemArray[1].ToString(), new string[] { @"mm\:ss", @"m\:ss" }, null, out TimeSpan ActionTime))
                    {
                        Data.Add(new DataStruct(ActionTime, new float[] {
                            float.Parse(Row.ItemArray[2].ToString()) ,
                            float.Parse(Row.ItemArray[3].ToString()) ,
                            float.Parse(Row.ItemArray[4].ToString()) ,
                            float.Parse(Row.ItemArray[5].ToString()),
                            float.Parse(Row.ItemArray[6].ToString()),
                            float.Parse(Row.ItemArray[7].ToString())
                        }));

                        //Debug.Log(ID + " ActionTime : " + ActionTime.TotalSeconds + " (" + ActionTime.Minutes.ToString("00") + ":" + ActionTime.Seconds.ToString("00") + ")");

                        //ID++;
                    }
                }

            }
        }

        DateTime StartTime = DateTime.Now;
        //for (int a = 0; a < Data.Count; a++)
        //{
        //    yield return new WaitUntil(() => StartTime + Data[a].ActionTime >= DateTime.Now);

        //    //Debug.Log(a + " : " + string.Join(", " + Data[a].FanArray));
        //    Debug.Log
        //}

        int ID = 0;
        while (true)
        {
            //if (StartTime + Data[ID].ActionTime >= DateTime.Now && !Data[ID].IsTrigger)
            //{
            //    Debug.Log()
            //}

            yield return null;
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

[Serializable]
struct DataStruct
{
    public TimeSpan ActionTime;
    public float[] FanArray;
    /// <summary>
    /// 是否已觸發
    /// </summary>
    public bool IsTrigger;

    public DataStruct(TimeSpan ActionTime, float[] FanArray)
    {
        this.ActionTime = ActionTime;
        this.FanArray = FanArray;
        IsTrigger = false;
    }
}