//! Indexer declarations: single-param, multi-param, get+set, read-only, interface-backed.
use super::helpers::run_csharp;

#[test]
fn single_parameter_indexer_reads_and_writes_backing_array() {
    assert_eq!(
        run_csharp(
            r#"class Vec{
    int[] data=new int[3];
    public int this[int i]{get=>data[i]; set=>data[i]=value;}
}
var v=new Vec(); v[1]=42;
Console.WriteLine(v[1]);"#
        ),
        &["42"]
    );
}

#[test]
fn two_parameter_indexer_simulates_matrix_access() {
    assert_eq!(
        run_csharp(
            r#"class Matrix{
    double[,] _d=new double[3,3];
    public double this[int r,int c]{get=>_d[r,c]; set=>_d[r,c]=value;}
}
var m=new Matrix(); m[1,2]=9.9;
Console.WriteLine(m[1,2]);"#
        ),
        &["9.9"]
    );
}

#[test]
fn readonly_indexer_exposes_computed_value() {
    assert_eq!(
        run_csharp(
            r#"class Odds{public int this[int n]=>2*n+1;}
Console.WriteLine(new Odds()[4]);"#
        ),
        &["9"]
    );
}

#[test]
fn string_indexed_dictionary_wrapper_exposes_indexer() {
    assert_eq!(
        run_csharp(
            r#"class Config{
    System.Collections.Generic.Dictionary<string,string> _d=new();
    public string this[string k]{get=>_d[k]; set=>_d[k]=value;}
}
var c=new Config(); c["env"]="prod";
Console.WriteLine(c["env"]);"#
        ),
        &["prod"]
    );
}

#[test]
fn interface_defines_indexer_contract_implemented_by_class() {
    assert_eq!(
        run_csharp(
            r#"interface IMap{string this[int k]{get;}}
class Map:IMap{string[] data={"zero","one","two"};public string this[int k]=>data[k];}
IMap m=new Map();
Console.WriteLine(m[2]);"#
        ),
        &["two"]
    );
}

#[test]
fn indexer_supports_negative_index_pattern_via_index_type() {
    assert_eq!(
        run_csharp(
            r#"int[] arr={1,2,3,4,5};
Console.WriteLine(arr[^1]); Console.WriteLine(arr[^2]);"#
        ),
        &["5", "4"]
    );
}
