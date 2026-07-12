//! Instance indexers (`this[key]`) expose element access distinct from fields.
use super::helpers::run_csharp;

#[test]
fn int_indexer_reads_written_slot() {
    assert_eq!(
        run_csharp(
            r#"
class Buffer {
    int[] data = new int[3];
    public int this[int index] {
        get { return data[index]; }
        set { data[index] = value; }
    }
}
var buffer = new Buffer();
buffer[1] = 42;
Console.WriteLine(buffer[1]);
"#
        ),
        &["42"]
    );
}

#[test]
fn string_keyed_indexer_stores_and_retrieves_values() {
    assert_eq!(
        run_csharp(
            r#"
class Bag {
    System.Collections.Generic.Dictionary<string, int> map = new();
    public int this[string key] {
        get { return map[key]; }
        set { map[key] = value; }
    }
}
var bag = new Bag();
bag["count"] = 7;
Console.WriteLine(bag["count"]);
"#
        ),
        &["7"]
    );
}

#[test]
fn indexer_get_can_compute_derived_value_from_state() {
    assert_eq!(
        run_csharp(
            r#"
class Scale {
    int factor = 2;
    public int this[int input] { get { return input * factor; } }
}
Console.WriteLine(new Scale()[5]);
"#
        ),
        &["10"]
    );
}

#[test]
fn indexer_set_can_update_multiple_fields_atomically() {
    assert_eq!(
        run_csharp(
            r#"
class PairStore {
    public int First;
    public int Second;
    public int this[int slot] {
        get { return slot == 0 ? First : Second; }
        set { if (slot == 0) First = value; else Second = value; }
    }
}
var store = new PairStore();
store[0] = 3;
store[1] = 9;
Console.WriteLine(store.First);
Console.WriteLine(store.Second);
"#
        ),
        &["3", "9"]
    );
}

#[test]
fn indexer_on_readonly_wrapper_exposes_underlying_element() {
    assert_eq!(
        run_csharp(
            r#"
class ReadWrapper {
    readonly int[] data = { 5, 6 };
    public int this[int i] { get { return data[i]; } }
}
Console.WriteLine(new ReadWrapper()[0]);
"#
        ),
        &["5"]
    );
}

#[test]
fn derived_class_indexer_can_call_base_indexer_via_cast() {
    assert_eq!(
        run_csharp(
            r#"
class Base {
    protected int[] data = { 1, 2 };
    public virtual int this[int i] { get { return data[i]; } }
}
class Derived : Base {
    public override int this[int i] { get { return base[i] + 10; } }
}
Base item = new Derived();
Console.WriteLine(item[1]);
"#
        ),
        &["12"]
    );
}

#[test]
fn interface_indexer_is_invoked_through_interface_typed_reference() {
    assert_eq!(
        run_csharp(
            r#"
interface ICell {
    string this[int index] { get; }
}
class Row : ICell {
    string[] cells = { "a", "b" };
    public string this[int index] { get { return cells[index]; } }
}
ICell row = new Row();
Console.WriteLine(row[1]);
"#
        ),
        &["b"]
    );
}

#[test]
fn multi_dimensional_style_dual_indexer_reads_matrix_cell() {
    assert_eq!(
        run_csharp(
            r#"
class Matrix {
    int[,] grid = { { 1, 2 }, { 3, 4 } };
    public int this[int row, int col] { get { return grid[row, col]; } }
}
Console.WriteLine(new Matrix()[1, 0]);
"#
        ),
        &["3"]
    );
}

#[test]
fn indexer_parameter_can_be_enum_typed_key() {
    assert_eq!(
        run_csharp(
            r#"
enum Axis { X, Y }
class Point {
  int[] values = { 4, 8 };
  public int this[Axis axis] { get { return values[(int)axis]; } }
}
Console.WriteLine(new Point()[Axis.Y]);
"#
        ),
        &["8"]
    );
}

#[test]
fn indexer_get_executes_side_effect_before_returning_value() {
    assert_eq!(
        run_csharp(
            r#"
class Logger {
    int hits = 0;
    public int this[int key] {
        get { hits++; return key; }
    }
}
var logger = new Logger();
Console.WriteLine(logger[5]);
Console.WriteLine(logger.hits);
"#
        ),
        &["5", "1"]
    );
}
