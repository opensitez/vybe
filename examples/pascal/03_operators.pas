program OperatorsDemo;

var
  A, B, C: Integer;
  X, Y: Real;
  Flag1, Flag2: Boolean;
begin
  A := 17;
  B := 5;

  Writeln('Arithmetic:');
  Writeln('A + B = ', A + B);
  Writeln('A - B = ', A - B);
  Writeln('A * B = ', A * B);
  Writeln('A / B = ', A / B:0:2);
  Writeln('A div B = ', A div B);
  Writeln('A mod B = ', A mod B);

  X := 2.5;
  Y := 4.0;
  Writeln('Real ops: X * Y = ', X * Y:0:2);

  Writeln('Bitwise:');
  Writeln('A shl 2 = ', A shl 2);
  Writeln('A shr 2 = ', A shr 2);

  Flag1 := True;
  Flag2 := False;
  Writeln('Logical:');
  Writeln('Flag1 and Flag2 = ', Flag1 and Flag2);
  Writeln('Flag1 or Flag2 = ', Flag1 or Flag2);
  Writeln('not Flag1 = ', not Flag1);
  Writeln('Flag1 xor Flag2 = ', Flag1 xor Flag2);

  Writeln('Relational:');
  Writeln('A > B = ', A > B);
  Writeln('A = B = ', A = B);
  Writeln('A <> B = ', A <> B);
end.
