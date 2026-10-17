// Permanent V6 tests included in preparation.rs::tests, using real plan fields.
#[test]
fn ranked_history_requires_permission_and_disable_survives_handoff() {
    let owner=PathBuf::from("parent");let key=PathBuf::from("streams");
    let mut plan=PreparationPlan::default();
    plan.prioritize_current(owner.clone(),std::slice::from_ref(&key),true);
    assert!(plan.priority_history.is_empty());
    plan.set_retirement_admission(true);
    plan.prioritize_current(owner.clone(),std::slice::from_ref(&key),true);
    assert_eq!(plan.current_priority(),[key.clone()]);
    plan.set_retirement_admission(false);
    assert!(plan.priority_owner.is_none());assert!(plan.priority_history.is_empty());
    let mut incoming=PreparationPlan::default();incoming.set_retirement_admission(true);
    incoming.prioritize_current(owner,std::slice::from_ref(&key),true);
    incoming.enqueue(key,admission_test_language());plan.extend(incoming);
    assert_eq!(plan.retirement_admission,Some(false));
    assert!(plan.priority_history.is_empty());assert!(plan.current_priority().is_empty());
    assert_eq!(plan.len(),1,"opt-out preserves normal pending preparation");
}

#[test]
fn nested_drained_parent_restores_retained_siblings_without_reordering_admissions() {
    let mut plan=PreparationPlan::default();plan.set_retirement_admission(true);
    let parent=PathBuf::from("parent");let child=PathBuf::from("child");
    let siblings=[PathBuf::from("streams"),PathBuf::from("curl")];
    plan.admissions=VecDeque::from([PathBuf::from("cold"),siblings[0].clone()]);
    let original=plan.admissions.clone();
    plan.prioritize_current(parent.clone(),&siblings,true);
    plan.prioritize_current(child,&[],true);assert!(plan.current_priority().is_empty());
    plan.prioritize_current(parent.clone(),&[],false);
    assert_eq!(plan.current_priority(),siblings);assert_eq!(plan.admissions,original);
    plan.prioritize_current(parent.clone(),&[siblings[0].clone(),PathBuf::from("proxy")],false);
    assert_eq!(plan.current_priority(),[siblings[0].clone(),siblings[1].clone(),PathBuf::from("proxy")]);
    plan.prioritize_current(parent,&[],true);assert!(plan.current_priority().is_empty());
}

#[test]
fn explicit_full_empty_receiving_owner_wins_older_handoff() {
    let owner=PathBuf::from("parent");let mut receiving=PreparationPlan::default();
    receiving.set_retirement_admission(true);receiving.prioritize_current(owner.clone(),&[],true);
    let mut old=PreparationPlan::default();old.set_retirement_admission(true);
    old.prioritize_current(owner.clone(),&[PathBuf::from("stale")],true);
    receiving.extend(old);assert_eq!(receiving.priority_owner,Some(owner));
    assert!(receiving.current_priority().is_empty());
    let mut unset=PreparationPlan::default();let mut current=PreparationPlan::default();
    current.set_retirement_admission(true);current.prioritize_current(PathBuf::from("new"),&[PathBuf::from("near")],true);
    unset.extend(current);assert_eq!(unset.current_priority(),[PathBuf::from("near")]);
}

#[test]
fn ranked_history_bounds_total_hints_and_owners_including_empty_receipts() {
    let mut plan=PreparationPlan::default();plan.set_retirement_admission(true);
    for i in 0..400 {
        let owner=PathBuf::from(format!("owner{i}"));let key=PathBuf::from(format!("hint{i}"));
        plan.prioritize_current(owner.clone(),&[key.clone(),key],true);
        assert!(plan.priority_history.len()<=128);
        assert!(plan.priority_history.iter().map(|(_,keys)|keys.len()).sum::<usize>()<=128);
        assert!(plan.priority_history.iter().any(|(o,_)|o==&owner));
    }
    for i in 0..400 {plan.prioritize_current(PathBuf::from(format!("empty{i}")),&[],true);assert!(plan.priority_history.len()<=128);}
    let full:Vec<_>=(0..200).map(|i|PathBuf::from(format!("key{i}"))).collect();
    plan.prioritize_current(PathBuf::from("selected"),&full,true);
    assert_eq!(plan.current_priority().len(),128);
    assert_eq!(plan.priority_history.iter().map(|(_,keys)|keys.len()).sum::<usize>(),128);
}
