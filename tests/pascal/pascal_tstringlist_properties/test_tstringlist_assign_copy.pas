// vybe-test: pascal/pascal_tstringlist_properties/test_tstringlist_assign_copy
// origin: languages/pascal/tests/pascal/test_pascal_tstringlist_properties.rs
program Test;
{$mode delphi}
uses Classes;
var sl1, sl2: TStringList;
begin
  sl1 := TStringList.Create;
  sl1.Add('CopyLine1'); sl1.Add('CopyLine2');
  sl2 := TStringList.Create;
  sl2.Assign(sl1);
  WriteLn(sl2.Count);
  WriteLn(sl2[1]);
  sl1.Free; sl2.Free;
end.
