program ForInLoopDemo;

var
  Arr: array of Integer;
  S: string;
  Ch: Char;
  I: Integer;
begin
  SetLength(Arr, 5);
  Arr[0] := 10;
  Arr[1] := 20;
  Arr[2] := 30;
  Arr[3] := 40;
  Arr[4] := 50;

  Writeln('For-in over array:');
  for I in Arr do
    Write(I, ' ');
  Writeln;

  S := 'Pascal';
  Writeln('For-in over string:');
  for Ch in S do
    Write(Ch, ' ');
  Writeln;

  Writeln('For-in over range:');
  for I in [1, 3, 5, 7, 9] do
    Write(I, ' ');
  Writeln;
end.
