program InterfacesDemo;

type
  IPrintable = interface
    function ToText: string;
  end;

  IComparable = interface
    function CompareTo(Other: IComparable): Integer;
  end;

  TProduct = class(IPrintable, IComparable)
  private
    FName: string;
    FPrice: Real;
  public
    constructor Create(const AName: string; APrice: Real);
    function ToText: string;
    function CompareTo(Other: IComparable): Integer;
    property Name: string read FName;
    property Price: Real read FPrice;
  end;

constructor TProduct.Create(const AName: string; APrice: Real);
begin
  FName := AName;
  FPrice := APrice;
end;

function TProduct.ToText: string;
begin
  Result := Name + ': $' + FloatToStr(Price);
end;

function TProduct.CompareTo(Other: IComparable): Integer;
var
  OtherProduct: TProduct;
begin
  OtherProduct := Other as TProduct;
  if Price < OtherProduct.Price then
    Result := -1
  else if Price > OtherProduct.Price then
    Result := 1
  else
    Result := 0;
end;

procedure PrintItem(Item: IPrintable);
begin
  Writeln(Item.ToText);
end;

var
  P1, P2: TProduct;
begin
  P1 := TProduct.Create('Widget', 29.99);
  P2 := TProduct.Create('Gadget', 49.99);

  PrintItem(P1);
  PrintItem(P2);

  Writeln('Compare: ', P1.CompareTo(P2));

  P1.Free;
  P2.Free;
end.
