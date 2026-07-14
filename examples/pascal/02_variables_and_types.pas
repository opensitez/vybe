program VariablesAndTypes;

var
  Age: Integer;
  PiValue: Real;
  Name: string;
  IsActive: Boolean;
  Letter: Char;
  Count: Cardinal;
  Small: ShortInt;
  Big: Int64;
begin
  Age := 25;
  PiValue := 3.14159;
  Name := 'Delphi Developer';
  IsActive := True;
  Letter := 'A';
  Count := 1000;
  Small := -128;
  Big := 9007199254740992;

  Writeln('Age: ', Age);
  Writeln('Pi: ', PiValue:0:5);
  Writeln('Name: ', Name);
  Writeln('Active: ', IsActive);
  Writeln('Letter: ', Letter);
  Writeln('Count: ', Count);
  Writeln('Small: ', Small);
  Writeln('Big: ', Big);
end.
