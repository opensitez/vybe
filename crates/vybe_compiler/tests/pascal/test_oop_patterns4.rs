use super::helpers::run_pascal;

#[test]
fn test_oop_observer_simple() {
    let src = r#"
program T;
type
  TObserver = class
    procedure OnUpdate(msg: string); virtual;
  end;
  TLogger = class(TObserver)
    procedure OnUpdate(msg: string); override;
  end;

procedure TObserver.OnUpdate(msg: string);
begin
  WriteLn('base:' + msg);
end;

procedure TLogger.OnUpdate(msg: string);
begin
  WriteLn('log:' + msg);
end;

var
  obs: TObserver;
  log: TLogger;
begin
  obs := TObserver.Create;
  log := TLogger.Create;
  obs.OnUpdate('event1');
  log.OnUpdate('event2');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["base:event1", "log:event2"]);
}

#[test]
fn test_oop_strategy_sort() {
    let src = r#"
program T;
type
  TSorter = class
    function Compare(a, b: Integer): Integer; virtual;
  end;
  TAscSorter = class(TSorter)
    function Compare(a, b: Integer): Integer; override;
  end;
  TDescSorter = class(TSorter)
    function Compare(a, b: Integer): Integer; override;
  end;

function TSorter.Compare(a, b: Integer): Integer;
begin
  Result := 0;
end;

function TAscSorter.Compare(a, b: Integer): Integer;
begin
  if a < b then Result := -1
  else if a > b then Result := 1
  else Result := 0;
end;

function TDescSorter.Compare(a, b: Integer): Integer;
begin
  if a > b then Result := -1
  else if a < b then Result := 1
  else Result := 0;
end;

var
  asc: TAscSorter;
  desc: TDescSorter;
begin
  asc := TAscSorter.Create;
  desc := TDescSorter.Create;
  WriteLn(asc.Compare(3, 5));
  WriteLn(desc.Compare(3, 5));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["-1", "1"]);
}

#[test]
fn test_oop_command_pattern() {
    let src = r#"
program T;
type
  TCommand = class
    procedure Execute; virtual; abstract;
  end;
  TPrintCmd = class(TCommand)
    FText: string;
    procedure Execute; override;
  end;

procedure TPrintCmd.Execute;
begin
  WriteLn(FText);
end;

var
  cmd: TPrintCmd;
begin
  cmd := TPrintCmd.Create;
  cmd.FText := 'Execute command!';
  cmd.Execute;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Execute command!"]);
}

#[test]
fn test_oop_decorator_basic() {
    let src = r#"
program T;
type
  TText = class
    function GetText: string; virtual;
  end;
  TBoldText = class(TText)
    FInner: TText;
    function GetText: string; override;
  end;

function TText.GetText: string;
begin
  Result := 'plain';
end;

function TBoldText.GetText: string;
begin
  Result := '**' + FInner.GetText + '**';
end;

var
  plain: TText;
  bold: TBoldText;
begin
  plain := TText.Create;
  bold := TBoldText.Create;
  bold.FInner := plain;
  WriteLn(bold.GetText);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["**plain**"]);
}

#[test]
fn test_oop_template_method() {
    let src = r#"
program T;
type
  TReport = class
    procedure Generate;
    procedure Header; virtual;
    procedure Body; virtual;
    procedure Footer; virtual;
  end;
  TSummary = class(TReport)
    procedure Header; override;
    procedure Body; override;
    procedure Footer; override;
  end;

procedure TReport.Generate;
begin
  Header;
  Body;
  Footer;
end;

procedure TReport.Header;
begin
  WriteLn('--- Report ---');
end;

procedure TReport.Body;
begin
  WriteLn('(empty body)');
end;

procedure TReport.Footer;
begin
  WriteLn('--- End ---');
end;

procedure TSummary.Header;
begin
  WriteLn('=== Summary ===');
end;

procedure TSummary.Body;
begin
  WriteLn('Total: 100');
end;

procedure TSummary.Footer;
begin
  WriteLn('=== Done ===');
end;

var
  s: TSummary;
begin
  s := TSummary.Create;
  s.Generate;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["=== Summary ===", "Total: 100", "=== Done ==="]);
}

#[test]
fn test_oop_iterator_pattern() {
    let src = r#"
program T;
type
  TRange = class
    FCurrent, FMax: Integer;
    procedure Init(max: Integer);
    function HasNext: Boolean;
    function Next: Integer;
  end;

procedure TRange.Init(max: Integer);
begin
  FCurrent := 1;
  FMax := max;
end;

function TRange.HasNext: Boolean;
begin
  Result := FCurrent <= FMax;
end;

function TRange.Next: Integer;
begin
  Result := FCurrent;
  FCurrent := FCurrent + 1;
end;

var
  it: TRange;
begin
  it := TRange.Create;
  it.Init(4);
  while it.HasNext do
    WriteLn(it.Next);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn test_oop_null_object_pattern() {
    let src = r#"
program T;
type
  TLogger = class
    procedure Log(msg: string); virtual;
  end;
  TNullLogger = class(TLogger)
    procedure Log(msg: string); override;
  end;

procedure TLogger.Log(msg: string);
begin
  WriteLn(msg);
end;

procedure TNullLogger.Log(msg: string);
begin
end;

procedure DoWork(log: TLogger);
begin
  log.Log('started');
  log.Log('done');
end;

var
  real: TLogger;
  null: TNullLogger;
begin
  real := TLogger.Create;
  null := TNullLogger.Create;
  DoWork(real);
  DoWork(null);
  WriteLn('finished');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["started", "done", "finished"]);
}

#[test]
fn test_oop_builder_pattern() {
    let src = r#"
program T;
type
  TQuery = class
    FTable: string;
    FWhere: string;
    FLimit: Integer;
    function SetTable(t: string): TQuery;
    function SetWhere(w: string): TQuery;
    function SetLimit(l: Integer): TQuery;
    function Build: string;
  end;

function TQuery.SetTable(t: string): TQuery;
begin
  FTable := t;
  Result := Self;
end;

function TQuery.SetWhere(w: string): TQuery;
begin
  FWhere := w;
  Result := Self;
end;

function TQuery.SetLimit(l: Integer): TQuery;
begin
  FLimit := l;
  Result := Self;
end;

function TQuery.Build: string;
begin
  Result := 'SELECT * FROM ' + FTable;
  if FWhere <> '' then
    Result := Result + ' WHERE ' + FWhere;
  if FLimit > 0 then
    Result := Result + ' LIMIT ' + IntToStr(FLimit);
end;

var
  q: TQuery;
begin
  q := TQuery.Create;
  q.SetTable('users');
  q.SetWhere('active=1');
  q.SetLimit(10);
  WriteLn(q.Build);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["SELECT * FROM users WHERE active=1 LIMIT 10"]);
}

#[test]
fn test_oop_proxy_pattern() {
    let src = r#"
program T;
type
  TService = class
    function Request(req: string): string; virtual;
  end;
  TProxy = class(TService)
    FReal: TService;
    FCallCount: Integer;
    function Request(req: string): string; override;
  end;

function TService.Request(req: string): string;
begin
  Result := 'response:' + req;
end;

function TProxy.Request(req: string): string;
begin
  FCallCount := FCallCount + 1;
  Result := FReal.Request(req);
end;

var
  svc: TService;
  proxy: TProxy;
begin
  svc := TService.Create;
  proxy := TProxy.Create;
  proxy.FReal := svc;
  proxy.FCallCount := 0;
  WriteLn(proxy.Request('ping'));
  WriteLn(proxy.Request('data'));
  WriteLn(proxy.FCallCount);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["response:ping", "response:data", "2"]);
}

#[test]
fn test_oop_abstract_method_two_impls() {
    let src = r#"
program T;
type
  TShape = class
    function Area: Integer; virtual; abstract;
    function Describe: string; virtual;
  end;
  TSquare = class(TShape)
    FSide: Integer;
    function Area: Integer; override;
    function Describe: string; override;
  end;
  TRect2 = class(TShape)
    FW, FH: Integer;
    function Area: Integer; override;
    function Describe: string; override;
  end;

function TShape.Describe: string;
begin
  Result := 'shape';
end;

function TSquare.Area: Integer;
begin
  Result := FSide * FSide;
end;

function TSquare.Describe: string;
begin
  Result := 'square(' + IntToStr(FSide) + ')';
end;

function TRect2.Area: Integer;
begin
  Result := FW * FH;
end;

function TRect2.Describe: string;
begin
  Result := 'rect(' + IntToStr(FW) + 'x' + IntToStr(FH) + ')';
end;

var
  sq: TSquare;
  rc: TRect2;
begin
  sq := TSquare.Create;
  sq.FSide := 5;
  rc := TRect2.Create;
  rc.FW := 4;
  rc.FH := 6;
  WriteLn(sq.Describe + '=' + IntToStr(sq.Area));
  WriteLn(rc.Describe + '=' + IntToStr(rc.Area));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["square(5)=25", "rect(4x6)=24"]);
}

#[test]
fn test_oop_chain_of_responsibility() {
    let src = r#"
program T;
type
  THandler = class
    FNext: THandler;
    FLevel: Integer;
    function Handle(level: Integer; msg: string): Boolean; virtual;
  end;

function THandler.Handle(level: Integer; msg: string): Boolean;
begin
  if level >= FLevel then begin
    WriteLn('[' + IntToStr(FLevel) + '] ' + msg);
    Result := true;
  end else if FNext <> nil then
    Result := FNext.Handle(level, msg)
  else
    Result := false;
end;

var
  h1, h2, h3: THandler;
begin
  h1 := THandler.Create; h1.FLevel := 1;
  h2 := THandler.Create; h2.FLevel := 3;
  h3 := THandler.Create; h3.FLevel := 5;
  h1.FNext := h2;
  h2.FNext := h3;
  h3.FNext := nil;
  h1.Handle(2, 'info');
  h1.Handle(4, 'warn');
  h1.Handle(6, 'error');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["[1] info", "[3] warn", "[5] error"]);
}

#[test]
fn test_oop_state_machine() {
    let src = r#"
program T;
type
  TTrafficLight = class
    FState: Integer;
    procedure Next;
    function Color: string;
  end;

procedure TTrafficLight.Next;
begin
  FState := (FState + 1) mod 3;
end;

function TTrafficLight.Color: string;
begin
  case FState of
    0: Result := 'Red';
    1: Result := 'Yellow';
    2: Result := 'Green';
    else Result := '?';
  end;
end;

var
  light: TTrafficLight;
  i: Integer;
begin
  light := TTrafficLight.Create;
  light.FState := 0;
  for i := 1 to 4 do begin
    WriteLn(light.Color);
    light.Next;
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Red", "Yellow", "Green", "Red"]);
}

#[test]
fn test_oop_singleton_pattern() {
    let src = r#"
program T;
type
  TApp = class
    FName: string;
    class function Instance: TApp;
  end;

var
  GInstance: TApp = nil;

class function TApp.Instance: TApp;
begin
  if GInstance = nil then
    GInstance := TApp.Create;
  Result := GInstance;
end;

begin
  TApp.Instance.FName := 'MyApp';
  WriteLn(TApp.Instance.FName);
  WriteLn(TApp.Instance = TApp.Instance);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["MyApp", "true"]);
}

#[test]
fn test_oop_visitor_simple() {
    let src = r#"
program T;
type
  TNode = class
    FValue: Integer;
    procedure Accept(var total: Integer);
  end;

procedure TNode.Accept(var total: Integer);
begin
  total := total + FValue;
end;

var
  nodes: array[0..2] of TNode;
  total: Integer;
  i: Integer;
begin
  for i := 0 to 2 do begin
    nodes[i] := TNode.Create;
    nodes[i].FValue := (i + 1) * 10;
  end;
  total := 0;
  for i := 0 to 2 do
    nodes[i].Accept(total);
  WriteLn(total);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_oop_prototype_clone() {
    let src = r#"
program T;
type
  TConfig = class
    FHost: string;
    FPort: Integer;
    function Clone: TConfig;
  end;

function TConfig.Clone: TConfig;
var
  c: TConfig;
begin
  c := TConfig.Create;
  c.FHost := FHost;
  c.FPort := FPort;
  Result := c;
end;

var
  cfg1, cfg2: TConfig;
begin
  cfg1 := TConfig.Create;
  cfg1.FHost := 'localhost';
  cfg1.FPort := 8080;
  cfg2 := cfg1.Clone;
  cfg2.FPort := 9090;
  WriteLn(cfg1.FHost + ':' + IntToStr(cfg1.FPort));
  WriteLn(cfg2.FHost + ':' + IntToStr(cfg2.FPort));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["localhost:8080", "localhost:9090"]);
}

#[test]
fn test_oop_facade_pattern() {
    let src = r#"
program T;
type
  TSubA = class
    procedure DoA; 
  end;
  TSubB = class
    procedure DoB;
  end;
  TFacade = class
    FA: TSubA;
    FB: TSubB;
    procedure Initialize;
    procedure Run;
  end;

procedure TSubA.DoA;
begin
  WriteLn('SubA done');
end;

procedure TSubB.DoB;
begin
  WriteLn('SubB done');
end;

procedure TFacade.Initialize;
begin
  FA := TSubA.Create;
  FB := TSubB.Create;
end;

procedure TFacade.Run;
begin
  FA.DoA;
  FB.DoB;
  WriteLn('Facade complete');
end;

var
  f: TFacade;
begin
  f := TFacade.Create;
  f.Initialize;
  f.Run;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["SubA done", "SubB done", "Facade complete"]);
}

#[test]
fn test_oop_composite_tree() {
    let src = r#"
program T;
type
  TNode = class
    FName: string;
    FChildren: array[0..1] of TNode;
    FChildCount: Integer;
    procedure Add(child: TNode);
    procedure Print(indent: Integer);
  end;

procedure TNode.Add(child: TNode);
begin
  FChildren[FChildCount] := child;
  FChildCount := FChildCount + 1;
end;

procedure TNode.Print(indent: Integer);
var
  i: Integer;
begin
  WriteLn(StringOfChar(' ', indent) + FName);
  for i := 0 to FChildCount - 1 do
    FChildren[i].Print(indent + 2);
end;

var
  root, c1, c2: TNode;
begin
  root := TNode.Create; root.FName := 'root'; root.FChildCount := 0;
  c1 := TNode.Create; c1.FName := 'child1'; c1.FChildCount := 0;
  c2 := TNode.Create; c2.FName := 'child2'; c2.FChildCount := 0;
  root.Add(c1);
  root.Add(c2);
  root.Print(0);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["root", "  child1", "  child2"]);
}

#[test]
fn test_oop_flyweight_chars() {
    let src = r#"
program T;
type
  TCharFlyweight = class
    FChar: Char;
    function Render: string;
  end;

function TCharFlyweight.Render: string;
begin
  Result := 'char:' + FChar;
end;

var
  pool: array[0..25] of TCharFlyweight;
  i: Integer;
begin
  for i := 0 to 25 do begin
    pool[i] := TCharFlyweight.Create;
    pool[i].FChar := Chr(Ord('a') + i);
  end;
  WriteLn(pool[0].Render);
  WriteLn(pool[4].Render);
  WriteLn(pool[25].Render);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["char:a", "char:e", "char:z"]);
}

#[test]
fn test_oop_polymorphic_array() {
    let src = r#"
program T;
type
  TAnimal = class
    function Sound: string; virtual;
  end;
  TDog = class(TAnimal)
    function Sound: string; override;
  end;
  TCat = class(TAnimal)
    function Sound: string; override;
  end;

function TAnimal.Sound: string;
begin
  Result := '...';
end;

function TDog.Sound: string;
begin
  Result := 'Woof';
end;

function TCat.Sound: string;
begin
  Result := 'Meow';
end;

var
  animals: array[0..2] of TAnimal;
  i: Integer;
begin
  animals[0] := TDog.Create;
  animals[1] := TCat.Create;
  animals[2] := TDog.Create;
  for i := 0 to 2 do
    WriteLn(animals[i].Sound);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Woof", "Meow", "Woof"]);
}

#[test]
fn test_oop_event_dispatcher() {
    let src = r#"
program T;
type
  TEventHandler = class
    FName: string;
    procedure Handle(event: string); virtual;
  end;

procedure TEventHandler.Handle(event: string);
begin
  WriteLn(FName + ' received:' + event);
end;

var
  handlers: array[0..1] of TEventHandler;
  i: Integer;
begin
  handlers[0] := TEventHandler.Create;
  handlers[0].FName := 'H1';
  handlers[1] := TEventHandler.Create;
  handlers[1].FName := 'H2';
  for i := 0 to 1 do
    handlers[i].Handle('click');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["H1 received:click", "H2 received:click"]);
}
