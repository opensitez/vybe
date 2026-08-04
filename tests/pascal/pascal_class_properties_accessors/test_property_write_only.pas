// vybe-test: pascal/pascal_class_properties_accessors/test_property_write_only
// origin: languages/pascal/tests/pascal/test_pascal_class_properties_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type TSecretStore = class
  private FSecret: String;
  private procedure SetSecret(s: String);
  public property Secret: String write SetSecret;
  public function IsSecretSet: Boolean;
end;
procedure TSecretStore.SetSecret(s: String); begin FSecret := s; end;
function TSecretStore.IsSecretSet: Boolean; begin Result := Length(FSecret) > 0; end;
var store: TSecretStore;
begin
  store := TSecretStore.Create;
  store.Secret := 'P@ssword';
  WriteLn(store.IsSecretSet);
  store.Free;
end.
