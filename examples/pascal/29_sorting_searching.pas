program SortingSearching;

var
  Numbers: array of Integer;
  Names: array of string;
  I, Target, FoundAt: Integer;
  Found: Boolean;

procedure BubbleSort(var Arr: array of Integer);
var
  I, J, Temp: Integer;
begin
  for I := 0 to High(Arr) - 1 do
    for J := 0 to High(Arr) - I - 1 do
      if Arr[J] > Arr[J + 1] then
      begin
        Temp := Arr[J];
        Arr[J] := Arr[J + 1];
        Arr[J + 1] := Temp;
      end;
end;

function LinearSearch(const Arr: array of Integer; Value: Integer): Integer;
var
  I: Integer;
begin
  Result := -1;
  for I := 0 to High(Arr) do
    if Arr[I] = Value then
    begin
      Result := I;
      Exit;
    end;
end;

function BinarySearch(const Arr: array of Integer; Value: Integer): Integer;
var
  Lo, Hi, Mid: Integer;
begin
  Lo := 0;
  Hi := High(Arr);
  Result := -1;
  while Lo <= Hi do
  begin
    Mid := (Lo + Hi) div 2;
    if Arr[Mid] = Value then
    begin
      Result := Mid;
      Exit;
    end
    else if Arr[Mid] < Value then
      Lo := Mid + 1
    else
      Hi := Mid - 1;
  end;
end;

begin
  SetLength(Numbers, 8);
  Numbers[0] := 64;
  Numbers[1] := 34;
  Numbers[2] := 25;
  Numbers[3] := 12;
  Numbers[4] := 22;
  Numbers[5] := 11;
  Numbers[6] := 90;
  Numbers[7] := 5;

  Writeln('Before sort:');
  for I := 0 to High(Numbers) do
    Write(Numbers[I], ' ');
  Writeln;

  BubbleSort(Numbers);

  Writeln('After sort:');
  for I := 0 to High(Numbers) do
    Write(Numbers[I], ' ');
  Writeln;

  Target := 22;
  FoundAt := LinearSearch(Numbers, Target);
  Writeln('Linear search for ', Target, ': index ', FoundAt);

  FoundAt := BinarySearch(Numbers, Target);
  Writeln('Binary search for ', Target, ': index ', FoundAt);

  SetLength(Names, 4);
  Names[0] := 'Charlie';
  Names[1] := 'Alice';
  Names[2] := 'Bob';
  Names[3] := 'Diana';

  Sort(Names);
  Writeln('Sorted names:');
  for I := 0 to High(Names) do
    Writeln(Names[I]);
end.
