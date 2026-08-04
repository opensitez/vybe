// vybe-test: pascal/pascal_class_rtti_typeinfo/test_rtti_read_published_property_value
// origin: languages/pascal/tests/pascal/test_pascal_class_rtti_typeinfo.rs
program Test;
{$mode delphi}
uses Rtti;
type TUser = class
  private FAge: Integer;
  public constructor Create(AAge: Integer);
  published property Age: Integer read FAge write FAge;
end;
constructor TUser.Create(AAge: Integer); begin FAge := AAge; end;
var ctx: TRttiContext;
    u: TUser;
    v: TValue;
begin
  u := TUser.Create(35);
  ctx := TRttiContext.Create;
  v := ctx.GetType(TUser).GetProperty('Age').GetValue(u);
  WriteLn(v.AsInteger);
  u.Free;
  ctx.Free;
end.
