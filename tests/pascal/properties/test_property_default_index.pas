// vybe-test: pascal/properties/test_property_default_index
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TList = class
  private
    FItems: array[0..2] of string;
    FCount: Integer;
    function GetItem(i: Integer): string;
    procedure SetItem(i: Integer; const v: string);
  public
    property Items[i: Integer]: string read GetItem write SetItem; default;
    property Count: Integer read FCount;
    procedure Add(s: string);
  end;

function TList.GetItem(i: Integer): string;
begin
  Result := FItems[i];
end;

procedure TList.SetItem(i: Integer; const v: string);
begin
  FItems[i] := v;
end;

procedure TList.Add(s: string);
begin
  FItems[FCount] := s;
  FCount := FCount + 1;
end;

var
  lst: TList;
begin
  lst := TList.Create;
  lst.Add('alpha');
  lst.Add('beta');
  lst.Add('gamma');
  WriteLn(lst[0]);
  WriteLn(lst[1]);
  WriteLn(lst[2]);
  WriteLn(lst.Count);
end.
