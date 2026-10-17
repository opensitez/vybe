use super::bundle_from_source;

#[test]
fn physical_php_alias_keeps_magic_constants_and_worker_identity_equal() {
    vybe_language_php::register();
    let language = vybe_compiler_language();
    let base = std::env::temp_dir().join(format!("vybe-physical-source-{}", uuid::Uuid::new_v4()));
    let physical = base.join("physical");
    std::fs::create_dir_all(&physical).unwrap();
    let alias = base.join("alias");
    std::os::unix::fs::symlink(&physical, &alias).unwrap();
    let source = "<?php return [__FILE__, __DIR__, '__FILE__'];";
    let file = physical.join("unit.php");
    std::fs::write(&file, source).unwrap();
    let canonical = file.canonicalize().unwrap();
    let from_alias = bundle_from_source(source, language, alias.join("unit.php"));
    let from_physical = bundle_from_source(source, language, &canonical);
    assert_eq!(from_alias.sources[0].path, canonical);
    assert_eq!(from_alias.sources[0].code, from_physical.sources[0].code);
    assert!(from_alias.sources[0].code.contains("'__FILE__'"), "quoted markers stay literal");
    let pool = crate::preparation::PreparationPool::new(1, 1).unwrap();
    let prepared = pool.submit(canonical, language).unwrap().wait().unwrap();
    assert!(prepared.module.is_ok());
    assert!(prepared.matches_frontend(language));
    assert!(prepared.matches_source(&from_alias.sources[0].path, &from_alias.sources[0].code),
        "a physical alias must not discard the canonical worker result");
    assert!(!prepared.matches_source(&from_alias.sources[0].path, "<?php return 0;"));
    std::fs::remove_dir_all(base).unwrap();
}

#[test]
fn virtual_php_and_other_language_paths_keep_the_supplied_spelling() {
    vybe_language_php::register();
    let php = vybe_compiler_language();
    let virtual_path = std::path::PathBuf::from("<eval>");
    assert_eq!(bundle_from_source("<?php return 1;", php, &virtual_path).sources[0].path, virtual_path);
    let base = std::env::temp_dir().join(format!("vybe-other-source-{}", uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&base).unwrap();
    let physical = base.join("unit.input");
    let alias = base.join("alias.input");
    std::fs::write(&physical, "source").unwrap();
    std::os::unix::fs::symlink(&physical, &alias).unwrap();
    let mut other = php;
    other.name = "identity-control";
    assert_eq!(bundle_from_source("source", other, &alias).sources[0].path, alias);
    std::fs::remove_dir_all(base).unwrap();
}

fn vybe_compiler_language() -> crate::languages::Language {
    crate::languages::find_by_name("php").unwrap()
}
