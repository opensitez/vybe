use super::*;
use super::super::{WastParser, component_interfaces, component_instances};
use vybe_ast::canon::{ComponentInterfaceDecl as Decl, TypeDecl};
use std::sync::Arc;

fn parsed(source: &str) -> Pair<'_, Rule> {
    WastParser::parse(Rule::component, source).unwrap().next().unwrap()
}

fn definitions(source: Pair<'_, Rule>) -> Vec<Pair<'_, Rule>> {
    source.into_inner().filter(|part| part.as_rule() == Rule::component_definition)
        .map(|part| part.into_inner().next().unwrap()).collect()
}

fn normalize(source: &str) -> (Pair<'_, Rule>, WastWalker) {
    let pair = parsed(source);
    let mut walker = WastWalker::default();
    component_interfaces::normalize(pair.clone(), &mut walker).unwrap();
    (pair, walker)
}

#[test]
fn recursive_members_empty_groups_and_module_holes_bind_source_names() {
    let source = r#"(component
      (core rec (type $s (struct)) (type $worker (func (param i32))))
      (core rec)
      (core type $module (module))
      (core type $array (array (mut (ref null $s)))))"#;
    let (pair, mut walker) = normalize(source);
    let mut count = 0;
    let mut names = HashMap::new();
    let mut binary_rows = Vec::new();
    for declaration in definitions(pair) {
        binary_rows.push(project_binary(row(&walker, &declaration).unwrap(), &mut names, &mut count).unwrap());
        walk(&mut walker, declaration).unwrap();
    }
    assert_eq!(walker.core_type_space, 4);
    assert_eq!(count, 4);
    assert_eq!(names, walker.core_type_index);
    assert_eq!(names["s"], 0);
    assert_eq!(names["worker"], 1);
    assert_eq!(names["module"], 2);
    assert_eq!(names["array"], 3);
    assert!(matches!(walker.core_type_kinds[1], CoreTypeKind::Worker(Scalar::I32)));
    assert!(matches!(walker.core_type_kinds[2], CoreTypeKind::Module));
    assert!(matches!(&binary_rows[1], Def::CoreType(ComponentCoreType::Rec(group)) if group.explicit && group.types.is_empty()));
}

#[test]
fn source_proposal_metadata_survives_normalization_and_projection() {
    let source = r#"(component
      (core rec
        (type $object (descriptor $desc) (struct (field (mut i8))))
        (type $desc (describes $object) (struct)))
      (core type $shared (shared (struct (field (ref null (shared any))))))
      (core type $f (func (param (ref (exact $object))) (result contref)))
      (core type $c (cont $f))
      (core type $module (module
        (type $f (func))
        (import "m" "f" (func (exact (type $f))))
        (export "memory" (memory i64 1 2 shared (pagesize 1)))
        (export "table" (table shared i64 1 2 (shared funcref)))
        (export "global" (global (shared mut i32)))
        (export "tag" (tag (type $f))))))"#;
    let (pair, walker) = normalize(source);
    let rows = definitions(pair).into_iter().map(|declaration|
        row(&walker, &declaration).unwrap().definition.clone()).collect::<Vec<_>>();
    let Def::CoreType(ComponentCoreType::Rec(group)) = &rows[0] else { panic!() };
    assert_eq!(group.types[0].descriptor, Some(1));
    assert_eq!(group.types[1].describes, Some(0));
    let Def::CoreType(ComponentCoreType::Rec(shared)) = &rows[1] else { panic!() };
    assert!(shared.types[0].shared);
    let Def::CoreType(ComponentCoreType::Rec(cont)) = &rows[3] else { panic!() };
    assert!(matches!(cont.types[0].composite, Composite::Cont(3)));
    let Def::CoreType(ComponentCoreType::Module(module)) = &rows[4] else { panic!() };
    assert_eq!(module.len(), 6);
    use vybe_ast::canon::{ComponentCoreModuleDecl as M, ComponentCoreExternType as E};
    assert!(matches!(&module[1], M::Import { ty: E::Func { type_index: 0, exact: true }, .. }));
    assert!(matches!(&module[2], M::Export { ty: E::Memory { limits, page_size_log2: Some(0) }, .. }
        if limits.index64 && limits.shared && limits.initial == 1 && limits.maximum == Some(2)));
    assert!(matches!(&module[3], M::Export { ty: E::Table { element, limits }, .. }
        if element.shared && limits.shared && limits.index64));
    assert!(matches!(&module[4], M::Export { ty: E::Global { mutable: true, shared: true, .. }, .. }));
    let Def::CoreType(ComponentCoreType::Rec(function)) = &rows[2] else { panic!() };
    assert!(matches!(&function.types[0].composite, Composite::Func { params, .. }
        if matches!(params.as_slice(), [Value::Ref(reference)] if reference.exact && !reference.nullable)));
}

#[test]
fn nested_component_core_alias_uses_declaration_time_lexical_scope() {
    let source = r#"(component
      (core type $worker (func (param i64)))
      (core type $module (module))
      (component
        (alias outer 1 $module (core type $m))
        (alias outer 1 $worker (core type $w))
        (core type $after (struct)))
      (core type $late (func)))"#;
    let (pair, mut walker) = normalize(source);
    let parent = definitions(pair);
    let module = row(&walker, &parent[1]).unwrap().slots[0].definition.clone();
    let child = parent.into_iter().find(|part| part.as_rule() == Rule::component).unwrap();
    let aliases = definitions(child.clone());
    assert!(Arc::ptr_eq(&row(&walker, &aliases[0]).unwrap().slots[0].definition, &module));
    for declaration in definitions(child) { walk(&mut walker, declaration).unwrap(); }
    assert_eq!(walker.core_type_index["m"], 0);
    assert_eq!(walker.core_type_index["w"], 1);
    assert_eq!(walker.core_type_index["after"], 2);
    assert!(matches!(walker.core_type_kinds[0], CoreTypeKind::Module));
    assert!(matches!(walker.core_type_kinds[1], CoreTypeKind::Worker(Scalar::I64)));
}

#[test]
fn instance_projection_keeps_core_scopes_separate_from_component_types() {
    let source = r#"(component
      (core type $outer (func (param i32)))
      (type $i (instance
        (alias outer 1 $outer (core type $outer-alias))
        (core type $s (struct))
        (core type $a (array i16))
        (core type $module (module))
        (alias outer 0 $module (core type $self-alias))
        (type $f (func))
        (export "run" (func (type $f))))))"#;
    let (pair, mut walker) = normalize(source);
    let key = span(&pair);
    walker.component_binary_scope = Some(key);
    let retained = walker.component_interface_binaries[&key].clone();
    walker.comp_types.push(TypeDecl::Opaque("instance".into()));
    walker.comp_type_slots.push(0);
    component_instances::register(&mut walker, 0, 0).unwrap();
    component_instances::import(&mut walker, "api".into(), 0, None).unwrap();
    assert_eq!(walker.comp_func_space.len(), 1);
    assert_eq!(walker.comp_func_space[0].functype, Some(1));
    assert_eq!(walker.comp_types.len(), 2);
    assert!(Arc::ptr_eq(&retained, &walker.component_interface_binaries[&key]));
    let Def::InterfaceType { declarations, .. } = &retained.definitions[1] else { panic!() };
    assert_eq!(declarations.iter().filter(|d| matches!(d, Decl::CoreType(_) | Decl::AliasOuterCoreType { .. })).count(), 5);
    assert!(matches!(declarations[0], Decl::AliasOuterCoreType { count: 1, index: 0 }));
    assert!(matches!(declarations[4], Decl::AliasOuterCoreType { count: 0, index: 3 }));
}

#[test]
fn invalid_outer_indices_and_rec_references_are_not_filled_or_renumbered() {
    for source in [
        "(component (core type (func (param (ref 99)))))",
        "(component (core type (module)) (component (alias outer 1 1 (core type))))",
        "(component (type (instance (alias outer 1 0 (core type)))))",
        "(component (core rec (type $a (struct)) (type $a (struct))))",
    ] {
        let pair = parsed(source);
        let error = component_interfaces::normalize(pair, &mut WastWalker::default());
        assert!(error.is_err(), "invalid source accepted: {source}");
    }
}

#[test]
fn interface_alias_keeps_original_typed_declaration_and_raw_index() {
    let source = r#"(component
      (core type $m (module (export "mem" (memory i64 1 2 (pagesize 1)))))
      (type (instance (alias outer 1 $m (core type $alias)))))"#;
    let (pair, walker) = normalize(source);
    let key = span(&pair);
    let binary = &walker.component_interface_binaries[&key];
    let Def::InterfaceType { declarations, .. } = &binary.definitions[1] else { panic!() };
    let Def::CoreType(ComponentCoreType::Module(module)) = &binary.definitions[0] else { panic!() };
    use vybe_ast::canon::{ComponentCoreModuleDecl as M, ComponentCoreExternType as E};
    assert!(matches!(&module[0], M::Export { ty: E::Memory { page_size_log2: Some(0), limits }, .. }
        if limits.index64 && limits.initial == 1 && limits.maximum == Some(2)));
    assert_eq!(declarations.len(), 1);
    assert!(matches!(&declarations[0], Decl::AliasOuterCoreType { count: 1, index: 0 }));
}

#[test]
fn slots_retain_group_boundaries_and_alias_member_positions() {
    use vybe_ast::canon::ComponentCoreRecType;
    let ComponentCoreType::Rec(first) = ComponentCoreType::function(vec![], vec![]) else { panic!() };
    let mut scope = Scope::default();
    scope.define(ComponentCoreType::Rec(ComponentCoreRecType { explicit: true,
        types: vec![first.types[0].clone(), first.types[0].clone()] })).unwrap();
    assert!(scope.define(ComponentCoreType::Rec(ComponentCoreRecType { explicit: true, types: vec![] })).unwrap().is_empty());
    scope.define(ComponentCoreType::Module(vec![])).unwrap();
    scope.alias(&[], 0, 1).unwrap();
    assert_eq!(scope.len(), 4);
    assert_eq!(scope.slot(3).unwrap().member, Some(1));
    assert!(std::sync::Arc::ptr_eq(&scope.slot(0).unwrap().definition,
        &scope.slot(3).unwrap().definition));
    assert!(matches!(scope.slot(2).unwrap().definition.as_ref(), ComponentCoreType::Module(_)));
}

#[test]
fn module_alias_keeps_binary_anchor_instead_of_rebuilding_source_scopes() {
    let source = r#"(component
      (core type $f (func (param i64)))
      (core type $m (module (alias outer 1 $f (type $local))
        (export "f" (func (type $local)))))
      (component (alias outer 1 $m (core type $copy))
        (core type $other (func (param i32)))))"#;
    let (pair, walker) = normalize(source);
    let parent = definitions(pair);
    let original = &row(&walker, &parent[1]).unwrap().slots[0].definition;
    let child = parent.iter().find(|p| p.as_rule() == Rule::component).unwrap();
    let aliases = definitions(child.clone());
    let copied = row(&walker, &aliases[0]).unwrap();
    assert_eq!(copied.definition.normalized_outer_alias(), Some(vybe_ast::canon::ComponentOuterAlias {
        kind: vybe_ast::canon::ComponentOuterAliasKind::CoreType, count: 1, index: 1,
    }));
    assert!(Arc::ptr_eq(original, &copied.slots[0].definition));
    let ComponentCoreType::Module(module) = copied.slots[0].definition.as_ref() else { panic!() };
    use vybe_ast::canon::{ComponentCoreModuleDecl as M, ComponentCoreExternType as E};
    assert!(matches!(module[0], M::AliasOuterType { count: 1, index: 0 }));
    assert!(matches!(module[1], M::Export { ty: E::Func { type_index: 0, exact: false }, .. }));
}

#[test]
fn deep_outer_core_alias_keeps_source_member_and_origin() {
    let mut root = Scope::default();
    let ComponentCoreType::Rec(mut group) = ComponentCoreType::function(vec![], vec![]) else { panic!() };
    group.explicit = true;
    group.types.push(group.types[0].clone());
    root.define(ComponentCoreType::Rec(group)).unwrap();
    let middle = Scope::default();
    let mut leaf = Scope::default();
    let copied = leaf.alias(&[root.clone(), middle], 2, 1).unwrap();
    assert_eq!(copied.member, Some(1));
    assert!(Arc::ptr_eq(&copied.definition, &root.slot(1).unwrap().definition));
    assert!(leaf.alias(&[root.clone()], 2, 0).is_err());
    assert!(leaf.alias(&[root], 1, 2).is_err());
}

#[test]
fn defining_types_does_not_retain_prior_slot_prefixes() {
    let mut scope = Scope::default();
    let first = scope.define(ComponentCoreType::function(vec![], vec![])).unwrap().pop().unwrap();
    let weak = Arc::downgrade(&first.definition);
    for _ in 1..16384 { scope.define(ComponentCoreType::function(vec![], vec![])).unwrap(); }
    assert_eq!(scope.len(), 16384);
    assert_eq!(Arc::strong_count(&first.definition), 2);
    drop(scope);
    assert_eq!(Arc::strong_count(&first.definition), 1);
    drop(first);
    assert!(weak.upgrade().is_none());
}

#[test]
fn generic_and_compat_core_alias_declarations_keep_the_original_member_and_module_holes() {
    use vybe_ast::canon::{ComponentOuterAlias as Alias, ComponentOuterAliasKind as Kind};
    let ComponentCoreType::Rec(mut group) = ComponentCoreType::function(vec![], vec![]) else { panic!() };
    group.explicit = true;
    group.types.push(group.types[0].clone());
    let mut parent = Scope::default();
    parent.define(ComponentCoreType::Rec(group)).unwrap();
    parent.define(ComponentCoreType::Module(vec![])).unwrap();
    for constructor in [Def::CoreAliasOuterType {count:1,index:1},
        Def::AliasOuter(Alias {kind:Kind::CoreType,count:1,index:1})] {
        let mut scope=Scope::default();
        let member=scope.declaration(&constructor,&[parent.clone()]).unwrap().unwrap();
        assert_eq!(member[0].member,Some(1));
        assert!(Arc::ptr_eq(&member[0].definition,&parent.slot(1).unwrap().definition));
        let module=scope.declaration(&Def::AliasOuter(Alias {kind:Kind::CoreType,count:1,index:2}),&[parent.clone()])
            .unwrap().unwrap();
        assert_eq!(module[0].member,None);
        assert!(Arc::ptr_eq(&module[0].definition,&parent.slot(2).unwrap().definition));
        assert_eq!(scope.len(),2);
        for kind in [Kind::Type,Kind::CoreModule,Kind::Component] {
            assert!(scope.declaration(&Def::AliasOuter(Alias {kind,count:0,index:0}),&[]).unwrap().is_none());
        }
        assert_eq!(scope.len(),2,"other namespaces cannot consume core slots");
    }
}

#[test]
fn named_numeric_and_deep_generic_aliases_keep_typed_group_members() {
    use vybe_ast::canon::ComponentOuterAliasKind as Kind;
    let source=r#"(component
      (core rec (type $s (struct)) (type $worker (func (param i64))))
      (core type $module (module))
      (component
        (alias outer 1 $worker (core type $named))
        (alias outer 1 1 (core type $numeric))
        (alias outer 1 2 (core type $hole))
        (component (alias outer 2 $worker (core type $deep)))))"#;
    let (pair,walker)=normalize(source);
    let root=definitions(pair);
    let original=row(&walker,&root[0]).unwrap().slots[1].definition.clone();
    let child=root.into_iter().find(|p|p.as_rule()==Rule::component).unwrap();
    let declarations=definitions(child);
    for declaration in &declarations[..2] {
        let retained=row(&walker,declaration).unwrap();
        let alias=retained.definition.normalized_outer_alias().unwrap();
        assert_eq!((alias.kind,alias.count,alias.index),(Kind::CoreType,1,1));
        assert_eq!(retained.slots[0].member,Some(1));
        assert!(Arc::ptr_eq(&retained.slots[0].definition,&original));
        let mut count=0;let mut names=HashMap::new();
        assert_eq!(project_binary(retained,&mut names,&mut count).unwrap(),retained.definition);
        assert_eq!(count,1);
    }
    assert_eq!(row(&walker,&declarations[2]).unwrap().slots[0].member,None);
    let deep=definitions(declarations[3].clone()).pop().unwrap();
    let retained=row(&walker,&deep).unwrap();
    assert_eq!(retained.definition.normalized_outer_alias().unwrap().count,2);
    assert_eq!(retained.slots[0].member,Some(1));
    assert!(Arc::ptr_eq(&retained.slots[0].definition,&original));
}

#[test]
fn attributed_import_export_views_and_generic_type_alias_slots_reach_the_instance_projector() {
    use vybe_ast::canon::{ComponentBinary, ComponentExternalName, ComponentNameAttribute as A,
        ComponentOuterAlias as Alias, ComponentOuterAliasKind as Kind, ComponentItemKind,
        ComponentInterfaceRef as Ref, ValSpec};
    let name=|value:&str|ComponentExternalName {name:value.into(),attributes:vec![A::ExternalId(format!("identity:{value}"))]};
    let binary=ComponentBinary {definitions:vec![Def::ValueType(ValSpec::Prim("u8".into())),
        Def::NamedImport {name:name("byte-in"),ty:Ref::TypeEq(0)},
        Def::AliasOuter(Alias {kind:Kind::Type,count:0,index:1}),
        Def::NamedExport {name:name("byte-out"),kind:ComponentItemKind::Type,index:2,ty:None},
        Def::InterfaceType {is_component:false,declarations:vec![
            Decl::Type(Box::new(Def::FuncType {params:vec![],result:None,is_async:false})),
            Decl::NamedExport {name:name("run"),ty:Ref::Func(0)}]},
    ]};
    let bytes=vybe_platform_wasm::write_component_binary(&binary).unwrap();
    let binary=Arc::new(vybe_platform_wasm::read_component_binary(&bytes).unwrap());
    let mut walker=WastWalker::default();let key=(0,1);
    walker.component_binary_scope=Some(key);
    walker.component_interface_binaries.insert(key,binary.clone());
    walker.comp_types=vec![TypeDecl::Value(ValSpec::Prim("u8".into())),
        TypeDecl::Value(ValSpec::Ref(0)),TypeDecl::Value(ValSpec::Ref(1)),TypeDecl::Value(ValSpec::Ref(2)),
        TypeDecl::Opaque("instance".into())];
    walker.comp_type_slots=vec![0,1,2,3,4];
    component_instances::register(&mut walker,4,4).unwrap();
    component_instances::import(&mut walker,"api".into(),4,None).unwrap();
    assert_eq!(walker.comp_func_space.len(),1);
    assert_eq!(walker.comp_func_space[0].functype,Some(5));
    assert!(Arc::ptr_eq(&binary,&walker.component_interface_binaries[&key]));
    let Def::InterfaceType {declarations,..}=&binary.definitions[4] else {panic!()};
    let Decl::NamedExport {name,..}=&declarations[1] else {panic!()};
    assert_eq!(name.attributes,vec![A::ExternalId("identity:run".into())]);
}

#[test]
fn full_walker_retains_v2_binary_instead_of_scalar_projection() {
    let source = r#"(component
      (core rec (type $s (struct)) (type $a (array (ref null $s))))
      (core type $m (module (type $f (func)) (export "f" (func (type $f))))))"#;
    let module = super::super::parse(source).unwrap();
    let binary = module.canon.binary.expect("typed binary survives the complete walker");
    assert!(matches!(&binary.definitions[0], Def::CoreType(ComponentCoreType::Rec(group))
        if group.explicit && group.types.len() == 2));
    assert!(matches!(&binary.definitions[1], Def::CoreType(ComponentCoreType::Module(_))));
    let bytes = vybe_platform_wasm::write_component_binary(&binary).unwrap();
    assert_eq!(vybe_platform_wasm::read_component_binary(&bytes).unwrap(), *binary);
}

#[test]
fn nested_source_alias_counts_above_one_are_not_collapsed() {
    let source = r#"(component
      (core type $f (func (param i64)))
      (component (component
        (alias outer 2 $f (core type $copy)))))"#;
    let (pair, mut walker) = normalize(source);
    let middle = definitions(pair).into_iter().find(|p| p.as_rule() == Rule::component).unwrap();
    let leaf = definitions(middle).into_iter().find(|p| p.as_rule() == Rule::component).unwrap();
    let declaration = definitions(leaf).pop().unwrap();
    assert_eq!(row(&walker, &declaration).unwrap().definition.normalized_outer_alias(),
        Some(vybe_ast::canon::ComponentOuterAlias {kind: vybe_ast::canon::ComponentOuterAliasKind::CoreType, count: 2, index: 0}));
    walk(&mut walker, declaration).unwrap();
    assert!(matches!(walker.core_type_kinds[0], CoreTypeKind::Worker(Scalar::I64)));
    let module = super::super::parse(source).unwrap();
    assert!(module.canon.binary.is_some());
}

#[test]
#[ignore = "wat/wast 261 and 262 cannot encode core rec within interface type declarations"]
fn interface_recursive_group_text_encoder_boundary() {
    normalize("(component (type (instance (core rec (type (struct)) (type (array i16))))))");
}
