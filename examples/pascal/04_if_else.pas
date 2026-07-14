program IfElseDemo;

var
  Score: Integer;
  Grade: string;
begin
  Score := 85;

  if Score >= 90 then
    Grade := 'A'
  else if Score >= 80 then
    Grade := 'B'
  else if Score >= 70 then
    Grade := 'C'
  else if Score >= 60 then
    Grade := 'D'
  else
    Grade := 'F';

  Writeln('Score: ', Score, ' Grade: ', Grade);

  if (Score >= 60) and (Score < 70) then
    Writeln('Barely passed')
  else if Score >= 70 then
    Writeln('Passed comfortably');

  if not (Score < 50) then
    Writeln('Did not fail');
end.
