// vybe-test: pascal/procedural_types/procedural_type_string_transform_callback
// origin: languages/pascal/tests/pascal/test_procedural_types.rs
program T;
{$mode delphi}
uses SysUtils; function Transform(s: String; fn: function(c: Char): Char): String; var i: Integer; begin Result:=''; for i:=1 to Length(s) do Result:=Result+fn(s[i]); end; begin WriteLn(Transform('ab', function(c: Char): Char begin Result:=UpCase(c); end)); end.
