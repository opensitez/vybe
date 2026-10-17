use vybe_ast::canon::{
    ComponentBinary, ComponentBinaryDef as D, ComponentInterfaceDecl as I,
    ComponentInterfaceRef as R, ComponentItemKind as K, ComponentItemRef as Item, ValSpec,
};

fn roundtrip(component: &ComponentBinary) {
    let bytes = crate::write_component_binary(component).unwrap();
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    assert_eq!(*component, crate::read_component_binary(&bytes).unwrap());
}

fn all_sorts() -> ComponentBinary {
    let item = |name: &str, kind, index| Item {
        name: name.into(),
        kind,
        index,
    };
    let mut component = ComponentBinary {
        definitions: vec![
            D::FuncType {
                params: vec![],
                result: None,
                is_async: false,
            },
            D::InterfaceType {
                is_component: false,
                declarations: vec![],
            },
            D::InterfaceType {
                is_component: true,
                declarations: vec![],
            },
            D::ImportFunc {
                name: "function".into(),
                type_index: 0,
            },
            D::Import {
                name: "instance".into(),
                ty: R::Instance(1),
            },
            D::Import {
                name: "component".into(),
                ty: R::Component(2),
            },
            D::Import {
                name: "resource".into(),
                ty: R::TypeSubResource,
            },
            D::Import {
                name: "value".into(),
                ty: R::Value(ValSpec::Prim("u32".into())),
            },
            D::CoreModule(b"\0asm\x01\0\0\0".to_vec()),
            D::InstanceFromExports(vec![
                item("module", K::Module, 0),
                item("function", K::Func, 0),
                item("instance", K::Instance, 0),
                item("component", K::Component, 0),
                item("resource", K::Type, 3),
                item("value", K::Value, 0),
            ]),
        ],
    };
    for (name, kind) in [
        ("module", K::Module),
        ("function", K::Func),
        ("instance", K::Instance),
        ("component", K::Component),
        ("resource", K::Type),
        ("value", K::Value),
    ] {
        component.definitions.push(D::AliasInstanceExport {
            instance: 1,
            name: name.into(),
            kind,
        });
    }
    component.definitions.push(D::ValueType(ValSpec::Own(4)));
    component.definitions.push(D::FuncType {
        params: vec![("handle".into(), ValSpec::Ref(5))],
        result: None,
        is_async: false,
    });
    component.definitions.push(D::ImportFunc {
        name: "consume".into(),
        type_index: 6,
    });
    component.definitions.push(D::InstanceFromExports(vec![
        item("module", K::Module, 1),
        item("function", K::Func, 1),
        item("instance", K::Instance, 2),
        item("component", K::Component, 1),
        item("resource", K::Type, 4),
        item("value", K::Value, 1),
    ]));
    component.definitions.push(D::ExportInstance {
        name: "result".into(),
        index: 3,
    });
    component
}

#[test]
fn aliases_preserve_all_sorts_and_separate_index_spaces() {
    roundtrip(&all_sorts());
}

#[test]
fn alias_wrong_sort_and_missing_instance_are_rejected() {
    let mut component = all_sorts();
    let D::AliasInstanceExport { kind, .. } = &mut component.definitions[10] else {
        unreachable!()
    };
    *kind = K::Func;
    assert!(crate::write_component_binary(&component).is_err());
    let mut component = all_sorts();
    let D::AliasInstanceExport { instance, .. } = &mut component.definitions[10] else {
        unreachable!()
    };
    *instance = 2;
    assert!(crate::write_component_binary(&component).is_err());
}

#[test]
fn interface_aliases_preserve_exported_type_identity() {
    let component = ComponentBinary {
        definitions: vec![D::InterfaceType {
            is_component: true,
            declarations: vec![
                I::Type(Box::new(D::InterfaceType {
                    is_component: false,
                    declarations: vec![
                        I::Type(Box::new(D::ValueType(ValSpec::Record(vec![(
                            "x".into(),
                            ValSpec::Prim("u32".into()),
                        )])))),
                        I::Export {
                            name: "item".into(),
                            ty: R::TypeEq(0),
                        },
                    ],
                })),
                I::Import {
                    name: "provider".into(),
                    ty: R::Instance(0),
                },
                I::AliasInstanceExport {
                    instance: 0,
                    name: "item".into(),
                    kind: K::Type,
                },
                I::Export {
                    name: "item".into(),
                    ty: R::TypeEq(1),
                },
            ],
        }],
    };
    roundtrip(&component);
}

#[test]
fn interface_instance_aliases_preserve_their_own_index_space() {
    let mut component = ComponentBinary {
        definitions: vec![D::InterfaceType {
            is_component: true,
            declarations: vec![
                I::Type(Box::new(D::InterfaceType {
                    is_component: false,
                    declarations: vec![],
                })),
                I::Type(Box::new(D::InterfaceType {
                    is_component: false,
                    declarations: vec![
                        I::AliasOuterType { count: 1, index: 0 },
                        I::Export { name: "child".into(), ty: R::Instance(0) },
                    ],
                })),
                I::Import { name: "provider".into(), ty: R::Instance(1) },
                I::AliasInstanceExport {
                    instance: 0,
                    name: "child".into(),
                    kind: K::Instance,
                },
                I::AliasInstanceExport {
                    instance: 0,
                    name: "child".into(),
                    kind: K::Instance,
                },
                I::Export { name: "empty".into(), ty: R::Instance(0) },
            ],
        }],
    };
    roundtrip(&component);
    let D::InterfaceType { declarations, .. } = &mut component.definitions[0] else {
        unreachable!()
    };
    let I::AliasInstanceExport { kind, .. } = &mut declarations[3] else {
        unreachable!()
    };
    *kind = K::Func;
    assert!(crate::write_component_binary(&component).is_err());
}

#[test]
fn adjacent_aliases_group_across_leb_boundaries_without_crossing_definitions() {
    for count in [1, 127, 128, 129, 1000] {
        let mut component = ComponentBinary { definitions: vec![
            D::FuncType { params: vec![], result: None, is_async: false },
            D::ImportFunc { name: "run".into(), type_index: 0 },
            D::InstanceFromExports(vec![Item { name: "run".into(), kind: K::Func, index: 0 }]),
        ] };
        for _ in 0..count {
            component.definitions.push(D::AliasInstanceExport {
                instance: 0, name: "run".into(), kind: K::Func,
            });
        }
        component.definitions.push(D::ExportFunc {
            name: "run".into(), index: count, type_index: None,
        });
        component.definitions.push(D::AliasInstanceExport {
            instance: 0, name: "run".into(), kind: K::Func,
        });
        roundtrip(&component);
        let bytes = crate::write_component_binary(&component).unwrap();
        let sections = wasmparser::Parser::new(0).parse_all(&bytes)
            .filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::ComponentAliasSection(section) => Some(section.count()),
                _ => None,
            }).collect::<Vec<_>>();
        assert_eq!(sections, vec![count, 1]);
    }
}

#[test]
fn grouped_alias_can_reference_instance_created_by_prior_alias() {
    let component = ComponentBinary { definitions: vec![
        D::FuncType { params: vec![], result: None, is_async: false },
        D::ImportFunc { name: "run".into(), type_index: 0 },
        D::InstanceFromExports(vec![Item { name: "run".into(), kind: K::Func, index: 0 }]),
        D::InstanceFromExports(vec![Item { name: "child".into(), kind: K::Instance, index: 0 }]),
        D::AliasInstanceExport { instance: 1, name: "child".into(), kind: K::Instance },
        D::AliasInstanceExport { instance: 2, name: "run".into(), kind: K::Func },
        D::ExportFunc { name: "run".into(), index: 1, type_index: None },
    ] };
    roundtrip(&component);
}
