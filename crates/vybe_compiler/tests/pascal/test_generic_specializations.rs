/// Generic record types and specialized generic methods beyond test_generics_interfaces.rs.
use super::helpers::run_pascal;

#[test]
fn generic_record_pair_fields() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPair<TKey, TValue> = record
    Key: TKey;
    Value: TValue;
  end;
var p: TPair<String, Integer>;
begin
  p.Key := 'count';
  p.Value := 3;
  WriteLn(p.Key);
  WriteLn(p.Value);
end."#
        ),
        &["count", "3"]
    );
}

#[test]
fn generic_procedure_swap_values() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Swap<T>(var a, b: T);
var t: T;
begin
  t := a; a := b; b := t;
end;
var x, y: Integer;
begin
  x := 1; y := 9;
  Swap<Integer>(x, y);
  WriteLn(x); WriteLn(y);
end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn generic_function_identity() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Id<T>(v: T): T;
begin Result := v; end;
begin
  WriteLn(Id<String>('same'));
  WriteLn(Id<Integer>(88));
end."#
        ),
        &["same", "88"]
    );
}

#[test]
fn generic_class_list_push_pop() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStack<T> = class
  private
    FItems: array of T;
  public
    procedure Push(v: T);
    function Pop: T;
  end;
procedure TStack<T>.Push(v: T);
begin
  SetLength(FItems, Length(FItems) + 1);
  FItems[High(FItems)] := v;
end;
function TStack<T>.Pop: T;
begin
  Result := FItems[High(FItems)];
  SetLength(FItems, Length(FItems) - 1);
end;
var s: TStack<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(10);
  s.Push(20);
  WriteLn(s.Pop);
  WriteLn(s.Pop);
end."#
        ),
        &["20", "10"]
    );
}

#[test]
fn generic_method_on_generic_class() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBox<T> = class
  public
    Value: T;
    function Get: T;
  end;
function TBox<T>.Get: T;
begin Result := Value; end;
var b: TBox<String>;
begin
  b := TBox<String>.Create;
  b.Value := 'data';
  WriteLn(b.Get);
end."#
        ),
        &["data"]
    );
}

#[test]
fn generic_constraint_class_type_with_inherited() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAnimal = class
  public
    function Name: String; virtual; abstract;
  end;
  TDog = class(TAnimal)
  public
    function Name: String; override;
  end;
function TDog.Name: String; begin Result := 'dog'; end;
function Speak<T: TAnimal>(a: T): String;
begin Result := a.Name; end;
var d: TDog;
begin
  d := TDog.Create;
  WriteLn(Speak<TDog>(d));
end."#
        ),
        &["dog"]
    );
}
