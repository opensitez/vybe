program WithStatementDemo;

type
  TPoint = record
    X, Y: Integer;
  end;

  TRect = record
    TopLeft, BottomRight: TPoint;
  end;

  TPerson = record
    FirstName, LastName: string;
    Age: Integer;
  end;

function Area(const R: TRect): Integer;
begin
  Result := (R.BottomRight.X - R.TopLeft.X) * (R.BottomRight.Y - R.TopLeft.Y);
end;

var
  P: TPoint;
  R: TRect;
  Person: TPerson;
begin
  with P do
  begin
    X := 10;
    Y := 20;
  end;
  Writeln('Point: ', P.X, ',', P.Y);

  with R do
  begin
    TopLeft.X := 0;
    TopLeft.Y := 0;
    BottomRight.X := 100;
    BottomRight.Y := 50;
  end;
  Writeln('Rect area = ', Area(R));

  with Person do
  begin
    FirstName := 'John';
    LastName := 'Doe';
    Age := 30;
  end;
  Writeln(Person.FirstName, ' ', Person.LastName, ' is ', Person.Age);

  with R.TopLeft do
  begin
    X := 5;
    Y := 5;
  end;
  Writeln('New TL: ', R.TopLeft.X, ',', R.TopLeft.Y);
end.
