//! Direct AST leaves must apply the same live read/place policy as PEG.

#[test]
fn superglobal_reads_and_places_preserve_complete_peg_ast() {
    for source in [
        "<?php echo count($_SESSION) === 0 ? 'ok' : 'fail';",
        "<?php function f(){return count($_SESSION)+1;} f();",
        "<?php $_GET=[1]; $_POST=$_GET; echo count($_POST);",
        "<?php function f(){$_GET=[1]; return $_GET[0]+1;} f();",
        "<?php $x=$_SESSION+$_GET;",
        "<?php function f(){return $_POST == $_GET;}",
        "<?php function take(&$x){} take($_SESSION);",
        "<?php $_SESSION['x']=1; echo $_SESSION['x'];",
        "<?php $f=function(){return $_ENV;}; echo count($f());",
    ] {
        let direct = vybe_language_php::parse(source).unwrap();
        let peg = vybe_language_php::parse_with_peg(source).unwrap();
        assert_eq!(format!("{direct:?}"), format!("{peg:?}"), "{source}");
    }
}

#[test]
fn reported_superglobal_regressions_preserve_complete_peg_ast() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/php");
    for relative in [
        "php_session_start_id_name/test_php_session_unset_clears_superglobal.php",
        "sessions/session_unset_clears_superglobal.php",
        "sessions/session_superglobal_is_array.php",
        "superglobals/env_superglobal_bound_from_process_environment.php",
        "superglobals/array_merge_post_overrides_get.php",
        "php_session_encode_decode_data/test_php_session_decode_empty_string_clears_data.php",
    ] {
        let source = std::fs::read_to_string(root.join(relative)).unwrap();
        let direct = vybe_language_php::parse(&source).unwrap();
        let peg = vybe_language_php::parse_with_peg(&source).unwrap();
        assert_eq!(format!("{direct:?}"), format!("{peg:?}"), "{relative}");
    }
}

#[test]
#[ignore = "full local PHP public-route/PEG AST and diagnostic corpus"]
fn public_php_corpus_preserves_peg_ast_and_diagnostics() {
    std::thread::Builder::new().stack_size(64 * 1024 * 1024).spawn(|| {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut pending = vec![root.join("tests/php")];
        let mut files = Vec::new();
        while let Some(directory) = pending.pop() {
            for entry in std::fs::read_dir(directory).unwrap() {
                let path = entry.unwrap().path();
                if path.is_dir() { pending.push(path); }
                else if path.extension().is_some_and(|extension| extension == "php") {
                    files.push(path);
                }
            }
        }
        assert!(!files.is_empty(), "PHP corpus is missing");
        for name in ["post.php", "functions.php", "class-wp-query.php",
            "formatting.php", "deprecated.php", "general-template.php"] {
            files.push(root.join("examples/webroot/wordpress/wp-includes").join(name));
        }
        files.sort();
        let mut accepted = 0;
        for path in &files {
            let source = std::fs::read_to_string(path).unwrap();
            let direct = vybe_language_php::parse(&source);
            let peg = vybe_language_php::parse_with_peg(&source);
            if direct.is_ok() { accepted += 1; }
            let direct = format!("{direct:?}");
            let peg = format!("{peg:?}");
            if direct != peg {
                let offset = direct.bytes().zip(peg.bytes()).position(|(a, b)| a != b)
                    .unwrap_or_else(|| direct.len().min(peg.len()));
                panic!("public/PEG AST or diagnostic mismatch: {} at byte {offset}", path.display());
            }
        }
        eprintln!("public_php_corpus cases={} accepted={accepted}", files.len());
    }).unwrap().join().unwrap();
}
