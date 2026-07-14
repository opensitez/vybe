program InheritanceOverride;

type
  TShape = class
  protected
    FColor: string;
  public
    constructor Create(const AColor: string);
    function Area: Real; virtual; abstract;
    function Describe: string; virtual;
    property Color: string read FColor;
  end;

  TRectangle = class(TShape)
  private
    FWidth, FHeight: Real;
  public
    constructor Create(const AColor: string; W, H: Real);
    function Area: Real; override;
    function Describe: string; override;
  end;

  TCircle = class(TShape)
  private
    FRadius: Real;
  public
    constructor Create(const AColor: string; R: Real);
    function Area: Real; override;
    function Describe: string; override;
  end;

constructor TShape.Create(const AColor: string);
begin
  FColor := AColor;
end;

function TShape.Describe: string;
begin
  Result := 'A shape colored ' + Color;
end;

constructor TRectangle.Create(const AColor: string; W, H: Real);
begin
  inherited Create(AColor);
  FWidth := W;
  FHeight := H;
end;

function TRectangle.Area: Real;
begin
  Result := FWidth * FHeight;
end;

function TRectangle.Describe: string;
begin
  Result := 'Rectangle ' + FloatToStr(FWidth) + 'x' + FloatToStr(FHeight) +
            ' colored ' + Color;
end;

constructor TCircle.Create(const AColor: string; R: Real);
begin
  inherited Create(AColor);
  FRadius := R;
end;

function TCircle.Area: Real;
begin
  Result := Pi * FRadius * FRadius;
end;

function TCircle.Describe: string;
begin
  Result := 'Circle radius ' + FloatToStr(FRadius) + ' colored ' + Color;
end;

var
  Rect: TRectangle;
  Circ: TCircle;
begin
  Rect := TRectangle.Create('red', 10, 5);
  Circ := TCircle.Create('blue', 7);

  Writeln(Rect.Describe);
  Writeln('Area = ', Rect.Area:0:2);

  Writeln(Circ.Describe);
  Writeln('Area = ', Circ.Area:0:2);

  Rect.Free;
  Circ.Free;
end.
