lua_print! {
    test_utf8_codes_valid => { "local s=''; for p, c in utf8.codes('A') do s=s..p..':'..c end; print(s)", "1:65" },
    test_utf8_codes_multibyte => { "local s=''; for p, c in utf8.codes('你好') do s=s..p..':'..c..',' end; print(s)", "1:20320,4:22909," },
    test_utf8_codes_invalid => { "local ok, err = pcall(function() for p, c in utf8.codes('a\\xFFb') do end end); print(tostring(ok))", "false" },
    test_utf8_match_charpattern => { "local s=''; for c in string.gmatch('A你B', utf8.charpattern) do s=s..c end; print(s)", "A你B" },
    test_utf8_gsub_charpattern => { "local s = string.gsub('A你B', utf8.charpattern, 'X'); print(s)", "XXX" }
}
