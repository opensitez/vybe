program PropertiesDemo;

type
  TBankAccount = class
  private
    FBalance: Real;
    FOwner: string;
    FActive: Boolean;
    procedure SetBalance(Value: Real);
    function GetStatus: string;
  public
    constructor Create(const AOwner: string; InitialBalance: Real);
    procedure Deposit(Amount: Real);
    procedure Withdraw(Amount: Real);
    property Balance: Real read FBalance write SetBalance;
    property Owner: string read FOwner;
    property Active: Boolean read FActive write FActive;
    property Status: string read GetStatus;
  end;

constructor TBankAccount.Create(const AOwner: string; InitialBalance: Real);
begin
  FOwner := AOwner;
  FBalance := InitialBalance;
  FActive := True;
end;

procedure TBankAccount.SetBalance(Value: Real);
begin
  if Value >= 0 then
    FBalance := Value;
end;

function TBankAccount.GetStatus: string;
begin
  if FActive then
    Result := FOwner + ' has $' + FloatToStr(FBalance)
  else
    Result := 'Account inactive';
end;

procedure TBankAccount.Deposit(Amount: Real);
begin
  if Active and (Amount > 0) then
    FBalance := FBalance + Amount;
end;

procedure TBankAccount.Withdraw(Amount: Real);
begin
  if Active and (Amount > 0) and (Amount <= FBalance) then
    FBalance := FBalance - Amount;
end;

var
  Account: TBankAccount;
begin
  Account := TBankAccount.Create('Alice', 1000);
  Writeln(Account.Status);

  Account.Deposit(500);
  Writeln('After deposit: ', Account.Balance:0:2);

  Account.Withdraw(200);
  Writeln('After withdrawal: ', Account.Balance:0:2);

  Account.Balance := 5000;
  Writeln('Set balance: ', Account.Balance:0:2);

  Account.Free;
end.
