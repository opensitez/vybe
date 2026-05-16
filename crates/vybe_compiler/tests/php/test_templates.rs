use std::fs;
use std::path::Path;

use super::helpers::{compile_ok, run_prints};
use vybe_compiler::bundle::{Bundle, EntryPoint, SourceFile};

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

#[test]
fn bundled_ixr_xml_declarations_compile() {
    let temp_root = std::env::temp_dir().join(format!(
        "vybex_php_ixr_bundle_{}",
        uuid::Uuid::new_v4()
    ));
    let ixr_dir = temp_root.join("IXR");
    fs::create_dir_all(&ixr_dir).expect("create IXR dir");

    let entry_path = temp_root.join("index.php");
    let class_ixr_path = temp_root.join("class-IXR.php");
    let request_path = ixr_dir.join("class-IXR-request.php");
    let server_path = ixr_dir.join("class-IXR-server.php");

    fs::write(
        &entry_path,
        "<?php\nrequire_once __DIR__ . '/class-IXR.php';\n",
    )
    .expect("write entry");
    fs::write(
        &class_ixr_path,
        "<?php\nrequire_once __DIR__ . '/IXR/class-IXR-server.php';\nrequire_once __DIR__ . '/IXR/class-IXR-request.php';\n",
    )
    .expect("write class-IXR");
    fs::write(
        &request_path,
        r#"<?php
class IXR_Request {
    public $xml;

    public function __construct($method, $args) {
        $this->xml = <<<EOD
<?xml version="1.0"?>
<methodCall>
EOD;
    }
}
"#,
    )
    .expect("write IXR request");
    fs::write(
        &server_path,
        r#"<?php
class IXR_Server {
    function output($xml) {
        $charset = function_exists('get_option') ? get_option('blog_charset') : '';
        if ($charset)
            $xml = '<?xml version="1.0" encoding="'.$charset.'"?>'."\n".$xml;
        else
            $xml = '<?xml version="1.0"?>'."\n".$xml;
        echo $xml;
    }
}
"#,
    )
    .expect("write IXR server");

    let language = vybe_compiler::languages::find_by_name("php").expect("php language");
    let bundle = Bundle {
        name: "index".to_string(),
        language,
        sources: vec![SourceFile {
            path: entry_path.clone(),
            code: fs::read_to_string(&entry_path).expect("read entry"),
        }],
        wasm_files: vec![],
        entry_point: EntryPoint::Auto,
    };

    let prepared = bundle.prepared_module();
    let _ = fs::remove_dir_all(&temp_root);
    prepared.expect("bundled IXR XML declarations should parse");
}