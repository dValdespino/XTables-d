import std.algorithm.searching;
import std.typecons;
import std.string;
import std.stdio;

import core.stdc.string;
import core.stdc.stdlib;

import xtables;
import xl;

alias BookHandle  = void*;
alias SheetHandle = void*;
alias Format      = void*;
alias Font        = void*;

template SetCell(T){
	void SetCell(SheetHandle handle, int row, int col, T value, Format fmt = null){
		Format _make_float2_format(Format fmt,bool percent=false){
			fmt = xlFormatSetNumFormatA(percent?NumFormat.NUMFORMAT_PERCENT_D2:NumFormat.NUMFORMAT_NUMBER_D2);

			return fmt;
		}

		//Write Strings
		static if(T.stringof == string.stringof){
			xlSheetWriteStrA(handle,row,col,value, fmt);
		}else
		//Write VAR values
		static if (T.stringof == VAR.stringof){
			bool isPercent=false;
			writeln(value);
			switch(value.Type){
				//Get the underlying string from it
				case ColumnType.String:
					handle.SetCell(row,col,value.AsString(), fmt);
					break;

				//Get as float
				case ColumnType.Percent:
					isPercent=true;
				case ColumnType.Float:
					handle.SetCell(row,col,value.AsFloat(), _make_float2_format(fmt, isPercent));
					break;

				//Get as int
				case ColumnType.Int:
					handle.SetCell(row,col,value.AsInt(), fmt);
					break;

				default:
					assert(0);
			}
		//Write Numbers
		}else{
			xlSheetWriteNumA(handle,row,col,value, T.stringof==float.stringof?_make_float2_format(fmt):fmt);
		}
	}
}

void MergeCells(SheetHandle sheet, int row_start, int row_end, int col_start, int col_end){
	xlSheetSetMergeA(sheet, row_start, row_end, col_start, col_end);
}

void SaveToFile(BookHandle book, string fname){
	xlBookSaveA(book, fname);
}

byte[] Dump(BookHandle book){
	char* buff;
	uint len;

	xlBookSaveRawA(book, &buff, &len);

	return cast(byte[])buff[0..len];
}

string dstring(char* str){
	return cast(string)fromStringz(str);
}

char* cstring(string str){
	return cast(char*)toStringz(str);
}

class XLSBook: XBook{
	private BookHandle Handle;

	this(string name, string fname){
		super(name);
		this.FileName=fname;
	}

	void WriteTable(XTable table, ref SheetHandle sheet, ref int row){
		int TitleBackgroundColor=Color.COLOR_GRAY40;
		int DataBackgroundColor=Color.COLOR_GRAY25;

		Font headerFont=this.Handle.xlBookAddFontA(null);
		headerFont.xlFontSetColorA(Color.COLOR_GRAY80);
		headerFont.xlFontSetNameA("Open Sans");
		headerFont.xlFontSetSizeA(12);
		headerFont.xlFontSetBoldA(true);

		Font colFont=this.Handle.xlBookAddFontA(headerFont);
		colFont.xlFontSetColorA(Color.COLOR_BLACK);
		colFont.xlFontSetSizeA(11);
		colFont.xlFontSetBoldA(false);

		Font dataFont=this.Handle.xlBookAddFontA(headerFont);
		dataFont.xlFontSetColorA(Color.COLOR_BLACK);
		dataFont.xlFontSetSizeA(11);
		dataFont.xlFontSetBoldA(false);

		Font darkDataFont=this.Handle.xlBookAddFontA(headerFont);
		darkDataFont.xlFontSetBoldA(true);

		Format headerFormat=xlBookAddFormatA(this.Handle, null);
		headerFormat.xlFormatSetFontA(headerFont);
		headerFormat.xlFormatSetAlignHA(AlignH.ALIGNH_CENTER);
		headerFormat.xlFormatSetAlignVA(AlignV.ALIGNV_CENTER);
		headerFormat.xlFormatSetFillPatternA(FillPattern.FILLPATTERN_SOLID);
		headerFormat.xlFormatSetPatternForegroundColorA(TitleBackgroundColor);
		headerFormat.xlFormatSetBorderA(BorderStyle.BORDERSTYLE_MEDIUM);

		Format colFormat=xlBookAddFormatA(this.Handle,headerFormat);
		colFormat.xlFormatSetBorderA(BorderStyle.BORDERSTYLE_NONE);
		colFormat.xlFormatSetFontA(colFont);

		Format dataFormat=xlBookAddFormatA(this.Handle,null);
		dataFormat.xlFormatSetFontA(dataFont);
		dataFormat.xlFormatSetPatternForegroundColorA(DataBackgroundColor);

		dataFormat.xlFormatSetFillPatternA(FillPattern.FILLPATTERN_SOLID);

		Format darkDataFormat=xlBookAddFormatA(this.Handle,dataFormat);
		darkDataFormat.xlFormatSetPatternForegroundColorA(TitleBackgroundColor);
		darkDataFormat.xlFormatSetFontA(darkDataFont);
		darkDataFormat.xlFormatSetAlignHA(AlignH.ALIGNH_CENTER);
		darkDataFormat.xlFormatSetAlignVA(AlignV.ALIGNV_CENTER);

		//Merge the cells over the table
		sheet.MergeCells(row, row, 1, cast(int)table.Columns.length);
		//Write the Table's title
		sheet.SetCell(row, 1, table.Name.toUpper(), headerFormat);

		//Skip the title
		row++;

		//Write the Column Headers
		for (int column_id=0;column_id<table.Columns.length;column_id++){
			sheet.SetCell(row,column_id+1,table.Columns[column_id].Name, column_id==0?darkDataFormat:colFormat);
		}

		//Skip the Column Names
		row++;

		//Write every entry's data
		for (int entry_id=0; entry_id<table.Entries.length; entry_id++){
			for (int column_id=0;column_id<table.Columns.length;column_id++){
				sheet.SetCell(row+entry_id, column_id+1, table.Entries[entry_id].Values[column_id], column_id==0?darkDataFormat:dataFormat);
			}
		}

		//Skip the table
		row+=table.Entries.length;

		//And then add an extra row
		row++;
	}

	void WriteTables(){
		SheetHandle[string] sheets;
		int[string] sheet_last_row;

		foreach (XTable table; Tables){
			SheetHandle sheet = table.SheetName in sheets;

			if (!sheet){
				sheet=xlBookAddSheetA(this.Handle, table.SheetName, null);
				sheets[table.SheetName]=sheet;
				sheet_last_row[table.SheetName]=2; //Start always on the second row
			}

			WriteTable(table, sheet, sheet_last_row[table.SheetName]);
		}

		writeln("Pages: ",sheets);
	}

	override byte[] Dump(){
		this.Handle=xlCreateXMLBookCA();
		this.WriteTables();

		return this.Handle.Dump();
	}
}

void main()
{
	XLSBook book=new XLSBook("Test Book", "testbook.xlsx");
	XTable table=new XTable("Test");

	table.Configure([
		"Id int optional",
		"Provincia string",
		"Municipio string",
		"Plan float",
		"Percent percent"
	]);

	table.SheetName="Test Book Sheet";

	book.AddTable(table);

	table.PushEntry(tuple(1,"Camaguey","Camaguey",666.666,12f));
	table.PushEntry(tuple(1,"Camaguey","IPVCE",1266.666,42f));
	table.PushEntry(tuple(1,"Camaguey","Minas",626.666,122f));
	table.PushEntry(tuple(1,"Camaguey","Najasa",66.666,1f));

	book.Save();

	writeln(table.Columns);
}
