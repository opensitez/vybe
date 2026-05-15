use std::fs;
use std::path::Path;

use super::helpers::{compile_ok, run_prints};

fn assert_rendered(src: &str, expected: &str) {
    assert_eq!(run_prints(src).join(""), expected);
}

#[test]
fn mixed_template_attribute_echo_runtime() {
    assert_rendered(
        r#"<?php $cols = [1, 2]; ?><td colspan="<?php echo count($cols)?>">ok</td>"#,
        r#"<td colspan="2">ok</td>"#,
    );
}

#[test]
fn mixed_template_inline_if_attribute_runtime() {
    assert_rendered(
        r#"<?php $w = 120; ?><td<?php if ($w>0) echo " style=\"width:".$w."px;\"";?>>x</td>"#,
        r#"<td style="width:120px;">x</td>"#,
    );
}

#[test]
fn mixed_template_adjacent_php_islands_runtime() {
    assert_rendered(
        r#"<?php $i = 1; $files = [null, ["isBack" => true]]; ?><tr class="snF <?php echo ($i%2==0) ? "snEven" : "snOdd"?><?php echo (isset($files[$i]["isBack"]) && $files[$i]["isBack"]) ? ' snBack' : '';?>"></tr>"#,
        r#"<tr class="snF snOdd snBack"></tr>"#,
    );
}

#[test]
fn webroot_example_compiles() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/webroot/index.php");
    let src = fs::read_to_string(path).expect("read webroot example");
    compile_ok(&src);
}