using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using UnityEngine.Profiling;
using Excel;
using System.IO;
using System.Data;

public class ReadExcel : MonoBehaviour{
	void Awake(){
		FileStream fs = File.Open (Application.dataPath + "/test.xlsx", FileMode.Open, FileAccess.Read);
		IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader (fs);
		do{
			Debug.Log(excelDataReader.Name);
			excelDataReader.Read();
			while(excelDataReader.Read()){
				for(int i=0;i<excelDataReader.FieldCount;i++){
					string value = excelDataReader.IsDBNull(i)?"":excelDataReader.GetString(i);
					Debug.Log(value);
				}
				
			}
		}while(excelDataReader.NextResult());

//		System.Object columns = result.Tables [0].Columns;s
//		System.Object rows = result.Tables [0].Rows;
//
//		for (int i = 0; i < rows; i++) {
//			for (int j = 0; j < columns; j++) {
//				string value = result.Tables [0].Rows [i] [j].ToString ();
//				Debug.Log (value);
//			}
//		}
	}
}
