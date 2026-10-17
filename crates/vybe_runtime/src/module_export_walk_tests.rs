use super::*;

// Rejected experiment: retain evidence without changing the hot direct lookup.
fn candidate<'a>(modules: &'a HashMap<String,ModuleRecord>, module: &str, name: &str) -> Option<&'a ExportEntry> {
    let mut module=module;
    let mut name=name;
    let mut export=modules.get(module)?.exports.get(name)?;
    if !matches!(export,ExportEntry::Indirect {..}) {return Some(export)}
    let mut visited=[None;8];
    let mut used=0;
    let mut overflow=Vec::new();
    loop {
        match export {
            ExportEntry::Indirect {from,name:target} => {
                let key=(module,name);
                if visited[..used].contains(&Some(key)) || overflow.contains(&key) {return None}
                if used<visited.len() {visited[used]=Some(key);used+=1} else {overflow.push(key)}
                module=from;name=target;
                export=modules.get(module)?.exports.get(name)?;
            }
            export => return Some(export),
        }
    }
}

fn previous<'a>(modules: &'a HashMap<String,ModuleRecord>, module: &str, name: &str) -> Option<&'a ExportEntry> {
    let mut module=module;
    let mut name=name;
    let mut visited=Vec::new();
    loop {
        let export=modules.get(module)?.exports.get(name)?;
        match export {
            ExportEntry::Indirect {from,name:target} => {
                if visited.contains(&(module,name)) {return None}
                visited.push((module,name));
                module=from;name=target;
            }
            export => return Some(export),
        }
    }
}

fn chain(depth: usize, cycle: Option<usize>) -> HashMap<String,ModuleRecord> {
    let mut modules=HashMap::new();
    for i in 0..=depth {
        let specifier=format!("module_{i}");
        let mut record=ModuleRecord::new_synthetic(specifier.clone());
        let export=if i<depth {
            ExportEntry::Indirect {from:format!("module_{}",i+1),name:format!("export_{}",i+1)}
        } else if let Some(target)=cycle {
            ExportEntry::Indirect {from:format!("module_{target}"),name:format!("export_{target}")}
        } else {ExportEntry::Function {idx:123}};
        record.exports.insert(format!("export_{i}"),export);
        modules.insert(specifier,record);
    }
    modules
}

#[test]
fn inline_and_spilled_walks_match_existing_resolution() {
    for depth in 0..=48 {
        for cycle in std::iter::once(None).chain((0..=depth).map(Some)) {
            let modules=chain(depth,cycle);
            for i in 0..=depth {
                let module=format!("module_{i}");
                let name=format!("export_{i}");
                let old=previous(&modules,&module,&name);
                let new=candidate(&modules,&module,&name);
                let production=resolve_host_export(&modules,&module,&name);
                assert_eq!(old.is_some(),new.is_some());
                assert_eq!(old.is_some(),production.is_some());
                if let (Some(old),Some(production))=(old,production) {assert!(std::ptr::eq(old,production))}
                if let (Some(old),Some(new))=(old,new) {assert!(std::ptr::eq(old,new))}
            }
        }
    }
    let mut modules=chain(12,None);
    assert!(resolve_host_export(&modules,"absent","export_0").is_none());
    assert!(resolve_host_export(&modules,"module_0","absent").is_none());
    modules.remove("module_9");
    assert!(resolve_host_export(&modules,"module_0","export_0").is_none());
}

#[test]
#[ignore = "paired export resolver measurement, not whole VM timing"]
fn export_walk_benchmark() {
    use std::{hint::black_box,time::Instant};
    for depth in [0usize,1,4,8,9,32] {
        let modules=chain(depth,None);
        let mut samples:[Vec<f64>;2]=std::array::from_fn(|_|Vec::new());
        for trial in 0..21 {
            for new in if trial%2==0 {[false,true]} else {[true,false]} {
                let begin=Instant::now();
                for _ in 0..4096 {
                    black_box(if new {candidate(black_box(&modules),black_box("module_0"),black_box("export_0"))}
                        else {previous(black_box(&modules),black_box("module_0"),black_box("export_0"))});
                }
                samples[usize::from(new)].push(begin.elapsed().as_secs_f64()/4096.);
            }
        }
        for sample in &mut samples {sample.sort_by(f64::total_cmp)}
        println!("depth={depth} old={:.3}ns inline={:.3}ns ratio={:.3}x",samples[0][10]*1e9,samples[1][10]*1e9,samples[0][10]/samples[1][10]);
    }
}
