module xtables;

import std.stdio;
import core.stdc.stdint;
import core.stdc.string;
import core.stdc.stdlib;
import std.string;

pragma(lib, "xl");

enum ColumnType{
    Int,
    String,
    Float,
    Percent
};

enum ColumnFlags{
    None,
    Optional,
    NoLoad
};

struct Column{
    string Name;
    ColumnType Type;
    ColumnFlags Flags = ColumnFlags.None;
    string Regexp;
    string[] Alternatives;
};

class XTable {
public:
    this(string name){
    	this.Name=name;
    }

    string Name;
    string SheetName="Hoja 1";

    int TableRow = -1;
    int TableColumn = -1;

    Column[] Columns;
    XTableEntry[] Entries;

    Column AddColumn(string name, ColumnType type, ColumnFlags flags = ColumnFlags.None){
    	Column ret=Column(name,type,flags,"",[]);
    	this.Columns~=ret;

    	return ret;
    }

    template PushEntry(T){
	    void PushEntry(T t){
	    	assert(t.length==this.Columns.length);
	    	assert(this.Entries.length==0 || this.Entries[$-1].length==this.Columns.length);

	    	foreach(v; t){
	    		this.Push(v);
	    	}
	    }
    }

	template Push(T){
	    void Push(T value){
            assert(this.Columns.length>0, "No columns defined for this table!");

	    	if (this.Entries.length==0 || this.Entries[$-1].length==this.Columns.length){
	    		this.Entries~=new XTableEntry();
	    	}

            ColumnType expected_type=this.Columns[this.Entries[$-1].length].Type;
            auto _var=VAR(value);
            _var.Type=expected_type;
	    	this.Entries[$-1].Values~=_var;

            writeln("PUSH: ", _var);
	    }
	}

    void AddTableEntry(XTableEntry entry){
    	assert (entry.length==this.Columns.length);

    	for(int i=0; i<entry.length; i++){
    		assert(entry.Values[i].Type==this.Columns[i].Type);
    	}

    	this.Entries~=entry;
    }

    void Configure(string[] fields){
    	ColumnFlags flags;

    	foreach(string f; fields){
    		flags=ColumnFlags.None;

    		string[] parts=split(f," ");
    		assert(parts.length>=2);

            string typestr=parts[1];
            string namestr=parts[0];
            string[] args=parts[2..$];

            import std.algorithm.searching;
    		if (args.length>0){
    			if (canFind(args,"optional")){
    				flags|=ColumnFlags.Optional;
    			}
    		}

    		switch(typestr){
    			case "int":
    				AddColumn(namestr, ColumnType.Int);
    				break;
    			case "string":
    				AddColumn(namestr, ColumnType.String);
    				break;
				case "float":
    				AddColumn(namestr, ColumnType.Float);
    				break;
                case "percent":
                    AddColumn(namestr, ColumnType.Percent);
                    break;
    			default:
    			assert(0);
    		}
    	}
    }
};

interface IBook{
    void Save(string fname="");
    byte[] Dump();

    void Load(string FileName);
    void Load(string buffer, size_t buffer_length);

    void Close();

    bool FindTable(XTable table, int page_id, out bool cancel);
}

abstract class XBook: IBook {
public:
    this(string name){
    	this.Name=name;
    }

    string Name;
    string FileName;

    XTable[] Tables;

    void AddTable(XTable table){
    	this.Tables~=table;
    }

    void Save(string fname=""){
    	if (fname=="")
    		fname=this.FileName;

    	import std.file;
    	std.file.write(FileName, Dump());
    	this.FileName=fname;
    }

    byte[] Dump(){
    	return [];
    }

    void Load(string FileName){}
    void Load(string buffer, size_t buffer_length){}

    void Close(){}

    bool FindTable(XTable table, int page_id, out bool cancel){return false;}
};

alias CPTR=void*;

import std.typecons;

template PushVarArgs(values){
	void Push(values){
		writeln(values);
	}
}

struct VAR{
	ColumnType Type=ColumnType.Int;
	CPTR Ptr;

	this(string str){
		this.Ptr=cast(void*)toStringz(str);
		this.Type=ColumnType.String;
	}

	this(int num){
		int* ptr=cast(int*)malloc(int.sizeof);
		*ptr=num;
		this.Ptr=cast(void*)ptr;
		this.Type=ColumnType.Int;
	}

	this(float num){
		float* ptr=cast(float*)malloc(float.sizeof);
		*ptr=num;
		this.Ptr=cast(void*)ptr;
		this.Type=ColumnType.Float;
	}

	string AsString(){
		return cast(string)fromStringz(cast(char*)Ptr);
	}

	int AsInt(){
		return *cast(int*)Ptr;
	}

	float AsFloat(){
		return *cast(float*)Ptr;
	}

	string toString(){
		string str=format("%s: ",this.Type);

		switch(this.Type){
			case ColumnType.Int:
				str~=format("'%d'",this.AsInt());
				break;
            case ColumnType.Float:
                str~=format("'~%.2f'",this.AsFloat());
                break;
            case ColumnType.Percent:
                str~=format("'~%.2f%%'",this.AsFloat());
                break;
			case ColumnType.String:
				str~="'"~this.AsString()~"'";
				break;
			default:
				str~="(Unknown)";
		}

		return str;
	}
}

class XTableEntry {
public:
    this(){}

    @property size_t length(){
    	return Values.length;
    }

    VAR[] Values;

    void Push(float value){
    	Values~=VAR(value);
    }

    void Push(int value){
    	Values~=VAR(value);
    }

    void Push(string value){
    	Values~=VAR(value);
    }

    override string toString(){
    	string str="{";

    	foreach(size_t i, VAR val;Values){
    		str~=val.toString();
    		if (i<length-1)
    			str~=", ";
    	}
    	str~="}";

    	return str;
    }
};