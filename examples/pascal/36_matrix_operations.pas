program MatrixOperations;

type
  TMatrix = array of array of Real;

function CreateMatrix(Rows, Cols: Integer): TMatrix;
var
  I, J: Integer;
begin
  SetLength(Result, Rows, Cols);
  for I := 0 to Rows - 1 do
    for J := 0 to Cols - 1 do
      Result[I, J] := I * Cols + J;
end;

procedure PrintMatrix(const M: TMatrix);
var
  I, J: Integer;
begin
  for I := 0 to High(M) do
  begin
    for J := 0 to High(M[0]) do
      Write(M[I, J]:6:1);
    Writeln;
  end;
end;

function AddMatrices(const A, B: TMatrix): TMatrix;
var
  I, J: Integer;
begin
  SetLength(Result, Length(A), Length(A[0]));
  for I := 0 to High(A) do
    for J := 0 to High(A[0]) do
      Result[I, J] := A[I, J] + B[I, J];
end;

function Transpose(const M: TMatrix): TMatrix;
var
  I, J: Integer;
begin
  SetLength(Result, Length(M[0]), Length(M));
  for I := 0 to High(M) do
    for J := 0 to High(M[0]) do
      Result[J, I] := M[I, J];
end;

function Trace(const M: TMatrix): Real;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to High(M) do
    Result := Result + M[I, I];
end;

var
  A, B, C, T: TMatrix;
begin
  A := CreateMatrix(3, 3);
  B := CreateMatrix(3, 3);

  Writeln('Matrix A:');
  PrintMatrix(A);

  Writeln('Matrix B:');
  PrintMatrix(B);

  C := AddMatrices(A, B);
  Writeln('A + B:');
  PrintMatrix(C);

  T := Transpose(A);
  Writeln('Transpose of A:');
  PrintMatrix(T);

  Writeln('Trace of A = ', Trace(A):0:1);
end.
