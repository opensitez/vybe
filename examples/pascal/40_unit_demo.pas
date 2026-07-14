program UnitDemo;

uses
  SysUtils;

function IsPalindrome(const S: string): Boolean;
var
  I, Len: Integer;
begin
  Len := Length(S);
  Result := True;
  for I := 1 to Len div 2 do
    if S[I] <> S[Len - I + 1] then
    begin
      Result := False;
      Exit;
    end;
end;

function CaesarCipher(const S: string; Shift: Integer): string;
var
  I: Integer;
  Ch: Char;
begin
  Result := '';
  for I := 1 to Length(S) do
  begin
    Ch := S[I];
    if (Ch >= 'a') and (Ch <= 'z') then
      Ch := Chr((Ord(Ch) - Ord('a') + Shift) mod 26 + Ord('a'))
    else if (Ch >= 'A') and (Ch <= 'Z') then
      Ch := Chr((Ord(Ch) - Ord('A') + Shift) mod 26 + Ord('A'));
    Result := Result + Ch;
  end;
end;

function Levenshtein(const S, T: string): Integer;
var
  D: array of array of Integer;
  I, J: Integer;
begin
  SetLength(D, Length(S) + 1, Length(T) + 1);
  for I := 0 to Length(S) do
    D[I, 0] := I;
  for J := 0 to Length(T) do
    D[0, J] := J;
  for I := 1 to Length(S) do
    for J := 1 to Length(T) do
      if S[I] = T[J] then
        D[I, J] := D[I - 1, J - 1]
      else
      begin
        D[I, J] := D[I - 1, J];
        if D[I, J - 1] < D[I, J] then
          D[I, J] := D[I, J - 1];
        if D[I - 1, J - 1] < D[I, J] then
          D[I, J] := D[I - 1, J - 1];
        D[I, J] := D[I, J] + 1;
      end;
  Result := D[Length(S), Length(T)];
end;

var
  S: string;
begin
  S := 'radar';
  Writeln('IsPalindrome(''', S, ''') = ', IsPalindrome(S));
  Writeln('IsPalindrome(''hello'') = ', IsPalindrome('hello'));

  Writeln('Caesar(''abc'', 3) = ', CaesarCipher('abc', 3));
  Writeln('Caesar(''xyz'', 3) = ', CaesarCipher('xyz', 3));

  Writeln('Levenshtein(''kitten'', ''sitting'') = ', Levenshtein('kitten', 'sitting'));
  Writeln('Levenshtein(''hello'', ''hello'') = ', Levenshtein('hello', 'hello'));
end.
