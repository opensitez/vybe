// vybe-test: pascal/pascal_class_properties_accessors/test_property_chained_modification
// origin: languages/pascal/tests/pascal/test_pascal_class_properties_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type TBox = class
  private FWidth, FHeight: Integer;
  public property Width: Integer read FWidth write FWidth;
  public property Height: Integer read FHeight write FHeight;
  public procedure Scale(factor: Integer);
end;
procedure TBox.Scale(factor: Integer);
begin
  Width := Width * factor;
  Height := Height * factor;
end;
var b: TBox;
begin
  b := TBox.Create;
  b.Width := 10; b.Height := 20;
  b.Scale(2);
  WriteLn(b.Width);
  WriteLn(b.Height);
  b.Free;
end.
