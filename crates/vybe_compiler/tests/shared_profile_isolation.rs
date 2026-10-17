use std::sync::Arc;
use vybe_compiler::profile::{parse_profile, parse_profile_shared};

#[path = "module_end_hook_isolation.rs"]
mod lifecycle;

#[path = "request_superglobal_bootstrap.rs"]
mod request_bootstrap;

#[test]
fn shared_cache_is_content_keyed_and_owned_overrides_remain_private() {
    let source = "[info]\nname='shared-test'\n[compiler]\nentry_point='main'\n";
    let shared = parse_profile_shared(source).unwrap();
    let same = parse_profile_shared(&source.to_string()).unwrap();
    assert!(Arc::ptr_eq(&shared, &same));
    let mut owned = parse_profile(source).unwrap();
    owned.entry_point = Some("private".into());
    assert_eq!(parse_profile_shared(source).unwrap().entry_point.as_deref(), Some("main"));
    let mut overridden = shared.clone();
    Arc::make_mut(&mut overridden).entry_point = Some("other".into());
    assert_eq!(shared.entry_point.as_deref(), Some("main"));
    assert_eq!(overridden.entry_point.as_deref(), Some("other"));
    let changed = parse_profile_shared(&source.replace("main", "changed")).unwrap();
    assert!(!Arc::ptr_eq(&shared, &changed));
    assert_eq!(changed.entry_point.as_deref(), Some("changed"));
}

#[test]
fn concurrent_compilers_share_definition_without_sharing_overrides() {
    let source = "[info]\nname='concurrent-test'\n[compiler]\nentry_point='main'\n";
    let shared = parse_profile_shared(source).unwrap();
    let workers: Vec<_> = (0..4).map(|i| {
        let shared = shared.clone();
        std::thread::spawn(move || {
            let mut local = parse_profile_shared(source).unwrap();
            assert!(Arc::ptr_eq(&local, &shared));
            Arc::make_mut(&mut local).entry_point = Some(format!("entry{i}"));
            assert_eq!(shared.entry_point.as_deref(), Some("main"));
        })
    }).collect();
    for worker in workers { worker.join().unwrap(); }
    assert_eq!(shared.entry_point.as_deref(), Some("main"));
}

#[test]
fn shared_profile_emits_the_same_chunks_as_owned_profile() {
    let source = vybe_language_php::profile_source();
    let module = vybe_language_php::parse("<?php $x = 1 + 2; echo $x;").unwrap();
    let owned = vybe_compiler::Compiler::with_profile(parse_profile(source).unwrap())
        .compile(&module).unwrap();
    let shared = vybe_compiler::Compiler::with_shared_profile(parse_profile_shared(source).unwrap())
        .compile(&module).unwrap();
    assert_eq!(owned.len(), shared.len());
    for (a,b) in owned.iter().zip(&shared) {
        assert_eq!(a.name, b.name);
        assert_eq!(a.code, b.code, "{}", a.name);
        assert_eq!(a.local_count, b.local_count, "{}", a.name);
        assert_eq!(a.imports.iter().map(|i| (&i.module,&i.name)).collect::<Vec<_>>(),
                   b.imports.iter().map(|i| (&i.module,&i.name)).collect::<Vec<_>>());
    }
}

#[cfg(target_os = "macos")]
#[test]
fn profile_setup_cpu_measurement() {
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    fn cpu() -> u64 {
        let mut time = [0i64; 2];
        assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
        time[0] as u64 * 1_000_000_000 + time[1] as u64
    }
    let source = vybe_language_php::profile_source();
    let profile = parse_profile_shared(source).unwrap();
    let expected = profile.builtins.len();
    for round in 0..12 {
        for shared in if round % 2 == 0 { [false, true] } else { [true, false] } {
            let start = cpu();
            for _ in 0..100 {
                if shared {
                    let p = parse_profile_shared(std::hint::black_box(source)).unwrap();
                    assert_eq!(p.builtins.len(), expected);
                    drop(std::hint::black_box(p));
                } else {
                    let p = parse_profile(std::hint::black_box(source)).unwrap();
                    assert_eq!(p.builtins.len(), expected);
                    drop(std::hint::black_box(p));
                }
            }
            println!("PROFILE_CPU {round} {} {}", shared as u8, cpu() - start);
        }
    }
}
