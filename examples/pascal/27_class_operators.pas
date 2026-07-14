program ClassOperatorsDemo;

type
  TComplex = record
    Re, Im: Real;
    class operator Add(const A, B: TComplex): TComplex;
    class operator Subtract(const A, B: TComplex): TComplex;
    class operator Multiply(const A, B: TComplex): TComplex;
    class operator Negative(const A: TComplex): TComplex;
    class operator Equal(const A, B: TComplex): Boolean;
    function ToString: string;
  end;

  TVector2 = record
    X, Y: Real;
    class operator Add(const A, B: TVector2): TVector2;
    class operator Multiply(const A: TVector2; Scalar: Real): TVector2;
    function Length: Real;
  end;

class operator TComplex.Add(const A, B: TComplex): TComplex;
begin
  Result.Re := A.Re + B.Re;
  Result.Im := A.Im + B.Im;
end;

class operator TComplex.Subtract(const A, B: TComplex): TComplex;
begin
  Result.Re := A.Re - B.Re;
  Result.Im := A.Im - B.Im;
end;

class operator TComplex.Multiply(const A, B: TComplex): TComplex;
begin
  Result.Re := A.Re * B.Re - A.Im * B.Im;
  Result.Im := A.Re * B.Im + A.Im * B.Re;
end;

class operator TComplex.Negative(const A: TComplex): TComplex;
begin
  Result.Re := -A.Re;
  Result.Im := -A.Im;
end;

class operator TComplex.Equal(const A, B: TComplex): Boolean;
begin
  Result := (A.Re = B.Re) and (A.Im = B.Im);
end;

function TComplex.ToString: string;
begin
  Result := FloatToStr(Re) + ' + ' + FloatToStr(Im) + 'i';
end;

class operator TVector2.Add(const A, B: TVector2): TVector2;
begin
  Result.X := A.X + B.X;
  Result.Y := A.Y + B.Y;
end;

class operator TVector2.Multiply(const A: TVector2; Scalar: Real): TVector2;
begin
  Result.X := A.X * Scalar;
  Result.Y := A.Y * Scalar;
end;

function TVector2.Length: Real;
begin
  Result := Sqrt(X * X + Y * Y);
end;

var
  C1, C2, C3: TComplex;
  V1, V2: TVector2;
begin
  C1.Re := 3; C1.Im := 2;
  C2.Re := 1; C2.Im := 7;
  C3 := C1 + C2;
  Writeln('C1 + C2 = ', C3.ToString);

  C3 := C1 * C2;
  Writeln('C1 * C2 = ', C3.ToString);

  C3 := -C1;
  Writeln('-C1 = ', C3.ToString);

  Writeln('C1 = C2? ', C1 = C2);

  V1.X := 3; V1.Y := 4;
  V2.X := 1; V2.Y := 2;
  Writeln('V1 length = ', V1.Length:0:2);
  Writeln('V1 + V2 = (', (V1 + V2).X:0:0, ',', (V1 + V2).Y:0:0, ')');
  Writeln('V1 * 2 = (', (V1 * 2).X:0:0, ',', (V1 * 2).Y:0:0, ')');
end.
