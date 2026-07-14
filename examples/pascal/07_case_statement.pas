program CaseStatementDemo;

var
  DayNum: Integer;
  Grade: Char;
  Color: string;
begin
  DayNum := 3;
  case DayNum of
    1: Writeln('Monday');
    2: Writeln('Tuesday');
    3: Writeln('Wednesday');
    4: Writeln('Thursday');
    5: Writeln('Friday');
    6, 7: Writeln('Weekend');
  else
    Writeln('Invalid day');
  end;

  Grade := 'B';
  case Grade of
    'A': Writeln('Excellent');
    'B': Writeln('Good');
    'C': Writeln('Average');
    'D', 'F': Writeln('Poor');
  end;

  Color := 'red';
  case Color of
    'red', 'green', 'blue': Writeln('Primary color');
    'yellow', 'cyan', 'magenta': Writeln('Secondary color');
  else
    Writeln('Unknown color');
  end;
end.
