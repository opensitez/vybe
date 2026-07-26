use super::helpers::run_prints;

#[test]
fn test_addslashes_stripslashes_quotes_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "O'Reilly \"Book\" \\ Path \0";
$escaped = addslashes($str);
$restored = stripslashes($escaped);
echo ($str === $restored) ? 'roundtrip_ok' : 'err', "\n";
"#
        ),
        vec!["roundtrip_ok"]
    );
}
