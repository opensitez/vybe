
crate::php_cases! {
    mb_substitute_character_set_get => {
        r#"<?php
$old = mb_substitute_character();
mb_substitute_character(0x3013);
echo mb_substitute_character() . "|";
mb_substitute_character("none");
echo mb_substitute_character() . "|";
mb_substitute_character($old);
echo "restored";
"#,
        ["12307|none|restored"]
    };
}
