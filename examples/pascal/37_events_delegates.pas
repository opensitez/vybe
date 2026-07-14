program EventsDelegates;

type
  TNotifyEvent = procedure(Sender: TObject) of object;
  TIntegerEvent = procedure(Sender: TObject; Value: Integer) of object;

  TButton = class
  private
    FCaption: string;
    FOnClick: TNotifyEvent;
  public
    constructor Create(const ACaption: string);
    procedure Click;
    property Caption: string read FCaption;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
  end;

  TCounter = class
  private
    FCount: Integer;
    FOnThreshold: TIntegerEvent;
  public
    constructor Create;
    procedure Increment;
    property Count: Integer read FCount;
    property OnThreshold: TIntegerEvent read FOnThreshold write FOnThreshold;
  end;

  TForm = class
  private
    FClickCount: Integer;
    procedure ButtonClicked(Sender: TObject);
    procedure ThresholdReached(Sender: TObject; Value: Integer);
  public
    constructor Create;
    procedure RunDemo;
  end;

constructor TButton.Create(const ACaption: string);
begin
  FCaption := ACaption;
end;

procedure TButton.Click;
begin
  if Assigned(FOnClick) then
    FOnClick(Self);
end;

constructor TCounter.Create;
begin
  FCount := 0;
end;

procedure TCounter.Increment;
begin
  FCount := FCount + 1;
  if (FCount >= 5) and Assigned(FOnThreshold) then
    FOnThreshold(Self, FCount);
end;

constructor TForm.Create;
begin
  FClickCount := 0;
end;

procedure TForm.ButtonClicked(Sender: TObject);
begin
  FClickCount := FClickCount + 1;
  Writeln('Button "', TButton(Sender).Caption, '" clicked! Total: ', FClickCount);
end;

procedure TForm.ThresholdReached(Sender: TObject; Value: Integer);
begin
  Writeln('Threshold reached! Count = ', Value);
end;

procedure TForm.RunDemo;
var
  Btn: TButton;
  Counter: TCounter;
  I: Integer;
begin
  Btn := TButton.Create('OK');
  Btn.OnClick := ButtonClicked;

  Counter := TCounter.Create;
  Counter.OnThreshold := ThresholdReached;

  for I := 1 to 3 do
    Btn.Click;

  for I := 1 to 7 do
    Counter.Increment;

  Btn.Free;
  Counter.Free;
end;

var
  Form: TForm;
begin
  Form := TForm.Create;
  Form.RunDemo;
  Form.Free;
end.
