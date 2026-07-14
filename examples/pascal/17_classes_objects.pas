program ClassesObjects;

type
  TAnimal = class
  private
    FName: string;
    FAge: Integer;
  public
    constructor Create(const AName: string; AAge: Integer);
    procedure Speak; virtual;
    function Describe: string;
    property Name: string read FName;
    property Age: Integer read FAge;
  end;

  TDog = class(TAnimal)
  public
    procedure Speak; override;
    function Fetch: string;
  end;

  TCat = class(TAnimal)
  public
    procedure Speak; override;
    function Climb: string;
  end;

constructor TAnimal.Create(const AName: string; AAge: Integer);
begin
  FName := AName;
  FAge := AAge;
end;

procedure TAnimal.Speak;
begin
  Writeln('Some generic sound');
end;

function TAnimal.Describe: string;
begin
  Result := Name + ' is ' + IntToStr(Age) + ' years old';
end;

procedure TDog.Speak;
begin
  Writeln('Woof!');
end;

function TDog.Fetch: string;
begin
  Result := Name + ' fetched the ball';
end;

procedure TCat.Speak;
begin
  Writeln('Meow!');
end;

function TCat.Climb: string;
begin
  Result := Name + ' climbed the tree';
end;

var
  Dog: TDog;
  Cat: TCat;
begin
  Dog := TDog.Create('Rex', 3);
  Cat := TCat.Create('Whiskers', 2);

  Writeln(Dog.Describe);
  Dog.Speak;
  Writeln(Dog.Fetch);

  Writeln(Cat.Describe);
  Cat.Speak;
  Writeln(Cat.Climb);

  Dog.Free;
  Cat.Free;
end.
