// vybe-test: pascal/advanced/property_read_write
// origin: languages/pascal/tests/pascal/test_advanced.rs
program T;
{$mode delphi}
uses SysUtils;
type TBox = class
  private FWidth: Integer;
  public
    constructor Create(W: Integer);
    property Width: Integer read FWidth write FWidth;
  end;
constructor TBox.Create(W: Integer); begin FWidth := W; end;
var b: TBox;
begin
  b := TBox.Create(10);
  WriteLn(b.Width);
  b.Width := 25;
  WriteLn(b.Width);
end.
