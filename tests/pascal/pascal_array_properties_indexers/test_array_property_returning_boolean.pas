// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_returning_boolean
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TBitFlags = class
  private FFlags: array[0..7] of Boolean;
  private function GetBit(i: Integer): Boolean;
  private procedure SetBit(i: Integer; b: Boolean);
  public property Bits[i: Integer]: Boolean read GetBit write SetBit; default;
end;
function TBitFlags.GetBit(i: Integer): Boolean; begin Result := FFlags[i]; end;
procedure TBitFlags.SetBit(i: Integer; b: Boolean); begin FFlags[i] := b; end;
var bf: TBitFlags;
begin
  bf := TBitFlags.Create;
  bf[3] := True;
  WriteLn(bf[0]);
  WriteLn(bf[3]);
  bf.Free;
end.
