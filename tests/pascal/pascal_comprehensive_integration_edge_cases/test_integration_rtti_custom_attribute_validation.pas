// vybe-test: pascal/pascal_comprehensive_integration_edge_cases/test_integration_rtti_custom_attribute_validation
// origin: languages/pascal/tests/pascal/test_pascal_comprehensive_integration_edge_cases.rs
program Test;
{$mode delphi}
uses Rtti;

type NonEmptyAttribute = class(TCustomAttribute);

type TInputModel = class
  private FCode: String;
  public
    [NonEmpty]
    property Code: String read FCode write FCode;
end;

function ValidateModel(obj: TObject): Boolean;
var ctx: TRttiContext; t: TRttiType; prop: TRttiProperty; attr: TCustomAttribute;
begin
  Result := True;
  ctx := TRttiContext.Create;
  t := ctx.GetType(obj.ClassType);
  for prop in t.GetProperties do
    for attr in prop.GetAttributes do
      if attr is NonEmptyAttribute then
        if prop.GetValue(obj).AsString = '' then Exit(False);
  ctx.Free;
end;

var m: TInputModel;
begin
  m := TInputModel.Create;
  m.Code := '';
  WriteLn(ValidateModel(m));
  m.Code := 'ValidCode';
  WriteLn(ValidateModel(m));
  m.Free;
end.
