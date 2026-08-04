// vybe-test: pascal/idioms/property_read_write_basic
// origin: languages/pascal/tests/pascal/test_idioms.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TPerson = class
  private
    FName: String;
  public
    constructor Create(name: String);
    property Name: String read FName write FName;
  end;

constructor TPerson.Create(name: String);
begin FName := name; end;

var p: TPerson;
begin
  p := TPerson.Create('Alice');
  WriteLn(p.Name);
  p.Name := 'Bob';
  WriteLn(p.Name);
end.
