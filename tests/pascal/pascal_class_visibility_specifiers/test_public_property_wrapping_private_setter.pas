// vybe-test: pascal/pascal_class_visibility_specifiers/test_public_property_wrapping_private_setter
// origin: languages/pascal/tests/pascal/test_pascal_class_visibility_specifiers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TScoreTracker = class
  private FScore: Integer;
  private procedure SetScore(v: Integer);
  public property Score: Integer read FScore write SetScore;
end;
procedure TScoreTracker.SetScore(v: Integer);
begin
  if v >= 0 then FScore := v;
end;
var st: TScoreTracker;
begin
  st := TScoreTracker.Create;
  st.Score := 50;
  WriteLn(st.Score);
  st.Free;
end.
