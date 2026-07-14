program RecordsDemo;

type
  TDate = record
    Day: Integer;
    Month: Integer;
    Year: Integer;
  end;

  TPerson = record
    Name: string;
    Age: Integer;
    BirthDate: TDate;
  end;

function MakeDate(D, M, Y: Integer): TDate;
begin
  Result.Day := D;
  Result.Month := M;
  Result.Year := Y;
end;

function PersonToString(const P: TPerson): string;
begin
  Result := P.Name + ' (' + IntToStr(P.Age) + ') born ' +
            IntToStr(P.BirthDate.Day) + '/' +
            IntToStr(P.BirthDate.Month) + '/' +
            IntToStr(P.BirthDate.Year);
end;

var
  Today: TDate;
  Alice: TPerson;
begin
  Today := MakeDate(17, 5, 2026);
  Writeln('Today: ', Today.Day, '/', Today.Month, '/', Today.Year);

  Alice.Name := 'Alice';
  Alice.Age := 30;
  Alice.BirthDate := MakeDate(15, 3, 1996);
  Writeln(PersonToString(Alice));

  with Alice do
  begin
    Writeln('Name: ', Name);
    Writeln('Age: ', Age);
  end;
end.
