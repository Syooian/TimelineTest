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
    Dictionary<short, DMXDataStruct> DMXData = new Dictionary<short, DMXDataStruct>();

    /// <summary>
    /// 
    /// </summary>
    [SerializeField]
    MediaPlayer AVProPlayer;

    /// <summary>
    /// �v���ɶ����
    /// </summary>
    [SerializeField]
    Text Txt_Time1;
    /// <summary>
    /// �v���ɶ���� (�`���)
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

                //Debug.LogWarning("��ƪ�ƶq : " + Result.Tables.Count);

                foreach (var Table in Result.Tables)
                {
                    //Debug.Log(Table);

                    DataRowCollection DataRow = Result.Tables[Table.ToString()].Rows;
                    //Debug.LogWarning("Rows : " + DataRow.Count);

                    short Universe = short.Parse(DataRow[0].ItemArray[1].ToString());
                    DMXDataStruct UniverseData;
                    if (DMXData.ContainsKey(Universe))//�p�G��Ƥ��w����Universe����ơA�����X�~��W�[
                        UniverseData = DMXData[Universe];
                    else
                        UniverseData = new DMXDataStruct(Universe);

                    //Debug.LogWarning("Universe : " + Universe);

                    #region �O�U��ƭn����̤���Ū��
                    //�p�GExcel�ɦ���z���b�A�᭱�S�������n���ȡA��ꤣ�ݭn�o�q�A������ItemArray�����קY�i

                    int DmxEnd = 0;
                    for (int a = 3; a < DataRow[0].ItemArray.Length; a++)
                    {
                        if (int.TryParse(DataRow[0].ItemArray[a].ToString(), out int DMXID))
                        {
                            //Debug.Log("DMX ID : " + DMXID);
                        }
                        else
                        {
                            DmxEnd = a;
                            break;
                        }
                    }
                    #endregion

                    #region Ū���ɶ��P�ƭ�
                    for (int a = 2; a < DataRow.Count; a++)//�C�@��
                    {
                        //Debug.LogWarning(string.Join(", ", DataRow[a].ItemArray));

                        //�`���
                        var ActionTime = (int)TimeSpan.FromSeconds(int.Parse(DataRow[a].ItemArray[0].ToString())).TotalSeconds;
                        //Debug.Log(ActionTime.TotalSeconds + " - " + ToDateTimeString(ActionTime));

                        //��X�{�����ɶ��b��ơA�S���N����New�@��
                        ActionTimeStruct Timeline;
                        if (UniverseData.ActionTime.ContainsKey(ActionTime))
                            Timeline = UniverseData.ActionTime[ActionTime];
                        else
                        {
                            Timeline = new ActionTimeStruct();
                            Timeline.DevicesValue = new float[512];
                        }

                        for (int b = 3; b < DmxEnd; b++)
                        {
                            //Debug.Log(a + "-" + b + " : " + DataRow[a].ItemArray[b]);
                            Timeline.DevicesValue[int.Parse(DataRow[0].ItemArray[b].ToString())] = float.Parse(DataRow[a].ItemArray[b].ToString());
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

    // Update is called once per frame
    void Update()
    {
        if (AVProPlayer.Control.IsPlaying())
        {
            var CurrentTime = AVProPlayer.Control.GetCurrentTime();

            Txt_Time1.text = ToDateTimeString(TimeSpan.FromSeconds(CurrentTime));

            Txt_Time2.text = ((int)CurrentTime).ToString();
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
    public float[] DevicesValue;
    /// <summary>
    /// �O�_�wĲ�o
    /// </summary>
    public bool IsTrigger;
}