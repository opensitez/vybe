program TypeCastingDemo;

var
  R: Real;
  I: Integer;
  S: string;
  B: Boolean;
  C: Char;
begin
  R := 3.14159;
  I := Trunc(R);
  Writeln('Trunc(3.14159) = ', I);

  I := 42;
  R := I;
  Writeln('Real(42) = ', R:0:2);

  S := '123';
  I := StrToInt(S);
  Writeln('StrToInt(''123'') = ', I);

  S := '3.14';
  R := StrToFloat(S);
  Writeln('StrToFloat(''3.14'') = ', R:0:2);

  I := 456;
  S := IntToStr(I);
  Writeln('IntToStr(456) = ', S);

  R := 2.718;
  S := FloatToStr(R);
  Writeln('FloatToStr(2.718) = ', S);

  B := True;
  S := BoolToStr(B);
  Writeln('BoolToStr(True) = ', S);

  S := '1';
  B := StrToBool(S);
  Writeln('StrToBool(''1'') = ', B);

  I := 65;
  C := Chr(I);
  Writeln('Chr(65) = ', C);

  C := 'Z';
  I := Ord(C);
  Writeln('Ord(''Z'') = ', I);

  S := 'Hello';
  I := Length(S);
  Writeln('Length(''Hello'') = ', I);
end.
