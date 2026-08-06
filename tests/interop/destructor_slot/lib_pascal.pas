// A Pascal unit linked as a SECONDARY unit. It deliberately carries no
// extraction header — it is not a test on its own, it is the other half of
// one, and `is_test_source` collects only files that have one.
//
// Do not name that header here, even to say it is absent: `is_test_source`
// matches the literal string anywhere in the first five lines, so a comment
// mentioning it makes this file register as a test of its own. It did.
program InteropLib;
{$mode delphi}

// The destructor RECORDS rather than prints. Whether a `WriteLn` from a
// secondary unit lands inside the entry unit's `ob_start()` buffer is a
// separate question about output buffering, and a test should turn on one
// thing at a time. Reading a value back asserts "the destructor ran" without
// depending on the answer.
var
  DestroyLog: String;

type
  TResource = class
  public
    Name: String;
    constructor Create(AName: String);
    destructor Destroy; override;
  end;

constructor TResource.Create(AName: String);
begin
  Name := AName;
end;

destructor TResource.Destroy;
begin
  DestroyLog := DestroyLog + 'destroyed:' + Name + ';';
end;

function MakeResource(AName: String): TResource;
begin
  Result := TResource.Create(AName);
end;

function GetDestroyLog: String;
begin
  Result := DestroyLog;
end;

begin
  DestroyLog := '';
end.
