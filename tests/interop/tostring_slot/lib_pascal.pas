// Secondary unit — carries no extraction header, so it is not collected as a
// test of its own. See the note in destructor_slot/lib_pascal.pas.
program InteropToStringLib;
{$mode delphi}
type
  TPoint = class
  public
    X, Y: Integer;
    constructor Create(AX, AY: Integer);
    function ToString: String; override;
  end;

constructor TPoint.Create(AX, AY: Integer);
begin
  X := AX;
  Y := AY;
end;

// Pascal spells the coercion role `ToString`. PHP spells it `__toString`.
// Neither one should have to know the other's spelling — both fill
// `ProtocolSlot::ToString`, declared in `languages/pascal/src/protocol.rs`
// and `languages/php/src/protocol.rs` respectively.
function TPoint.ToString: String;
begin
  Result := '(' + IntToStr(X) + ',' + IntToStr(Y) + ')';
end;

function MakePoint(AX, AY: Integer): TPoint;
begin
  Result := TPoint.Create(AX, AY);
end;

begin
end.
