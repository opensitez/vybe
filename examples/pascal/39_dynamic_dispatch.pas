program DynamicDispatchDemo;

type
  TAnimal = class
  public
    procedure Speak; virtual;
    function Species: string; virtual;
  end;

  TDog = class(TAnimal)
  public
    procedure Speak; override;
    function Species: string; override;
    function Breed: string; virtual;
  end;

  TCat = class(TAnimal)
  public
    procedure Speak; override;
    function Species: string; override;
    function Lives: Integer; virtual;
  end;

  TBird = class(TAnimal)
  public
    procedure Speak; override;
    function Species: string; override;
    function CanFly: Boolean; virtual;
  end;

procedure TAnimal.Speak;
begin
  Writeln('...');
end;

function TAnimal.Species: string;
begin
  Result := 'Unknown';
end;

procedure TDog.Speak;
begin
  Writeln('Woof!');
end;

function TDog.Species: string;
begin
  Result := 'Canis lupus';
end;

function TDog.Breed: string;
begin
  Result := 'Mixed';
end;

procedure TCat.Speak;
begin
  Writeln('Meow!');
end;

function TCat.Species: string;
begin
  Result := 'Felis catus';
end;

function TCat.Lives: Integer;
begin
  Result := 9;
end;

procedure TBird.Speak;
begin
  Writeln('Tweet!');
end;

function TBird.Species: string;
begin
  Result := 'Aves';
end;

function TBird.CanFly: Boolean;
begin
  Result := True;
end;

procedure MakeItSpeak(Animal: TAnimal);
begin
  Write(Animal.Species, ' says ');
  Animal.Speak;
end;

var
  Animals: array of TAnimal;
  I: Integer;
begin
  SetLength(Animals, 3);
  Animals[0] := TDog.Create;
  Animals[1] := TCat.Create;
  Animals[2] := TBird.Create;

  for I := 0 to High(Animals) do
    MakeItSpeak(Animals[I]);

  for I := 0 to High(Animals) do
    Animals[I].Free;
end.
