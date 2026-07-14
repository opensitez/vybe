program EnumsDemo;

type
  TColor = (Red, Green, Blue);
  TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
  TSize = (Small, Medium, Large);

function DayName(D: TDay): string;
begin
  case D of
    Mon: Result := 'Monday';
    Tue: Result := 'Tuesday';
    Wed: Result := 'Wednesday';
    Thu: Result := 'Thursday';
    Fri: Result := 'Friday';
    Sat: Result := 'Saturday';
    Sun: Result := 'Sunday';
  end;
end;

function IsWeekend(D: TDay): Boolean;
begin
  Result := (D = Sat) or (D = Sun);
end;

var
  C: TColor;
  Today: TDay;
  S: TSize;
begin
  C := Green;
  Writeln('Color ord: ', Ord(C));
  Writeln('Pred: ', Ord(Pred(C)));
  Writeln('Succ: ', Ord(Succ(C)));

  Today := Fri;
  Writeln('Today is ', DayName(Today));
  Writeln('Weekend? ', IsWeekend(Today));

  for S := Small to Large do
    Writeln('Size ord: ', Ord(S));

  Writeln('Low TDay: ', Ord(Low(TDay)));
  Writeln('High TDay: ', Ord(High(TDay)));
end.
