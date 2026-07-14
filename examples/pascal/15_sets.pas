program SetsDemo;

type
  TCharSet = set of Char;
  TNumSet = set of 0..20;
  TDaySet = set of (Mon, Tue, Wed, Thu, Fri, Sat, Sun);

procedure PrintNumSet(const S: TNumSet);
var
  I: Integer;
begin
  for I := 0 to 20 do
    if I in S then
      Write(I, ' ');
  Writeln;
end;

var
  Vowels: TCharSet;
  Evens, Primes, Union, Intersection: TNumSet;
  Weekend: TDaySet;
  C: Char;
begin
  Vowels := ['A', 'E', 'I', 'O', 'U'];
  C := 'E';
  Writeln(C, ' in Vowels? ', C in Vowels);
  Writeln('Z in Vowels? ', 'Z' in Vowels);

  Evens := [0, 2, 4, 6, 8, 10];
  Primes := [2, 3, 5, 7, 11, 13];
  Union := Evens + Primes;
  Intersection := Evens * Primes;

  Writeln('Evens:');
  PrintNumSet(Evens);
  Writeln('Primes:');
  PrintNumSet(Primes);
  Writeln('Union:');
  PrintNumSet(Union);
  Writeln('Intersection:');
  PrintNumSet(Intersection);

  Weekend := [Sat, Sun];
  Writeln('Fri in Weekend? ', Fri in Weekend);
  Writeln('Sat in Weekend? ', Sat in Weekend);
end.
