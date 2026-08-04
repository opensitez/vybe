// vybe-test: pascal/properties/test_property_with_notification
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TObservable = class
  private
    FValue: Integer;
    FChanges: Integer;
    procedure SetValue(v: Integer);
  public
    property Value: Integer read FValue write SetValue;
    property Changes: Integer read FChanges;
  end;

procedure TObservable.SetValue(v: Integer);
begin
  if v <> FValue then begin
    FValue := v;
    FChanges := FChanges + 1;
  end;
end;

var
  o: TObservable;
begin
  o := TObservable.Create;
  o.Value := 1;
  o.Value := 1;
  o.Value := 2;
  o.Value := 3;
  o.Value := 3;
  WriteLn(o.Value);
  WriteLn(o.Changes);
end.
