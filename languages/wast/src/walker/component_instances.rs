use super::{binary_valspec, validate_component_type_size, CompFunc, CompInstance, WastWalker};
use std::collections::HashMap;
use std::sync::Arc;
use vybe_ast::canon::{ComponentBinaryDef as Def, ComponentInterfaceDecl as Decl,
    ComponentInterfaceRef as Ref, ComponentBinary, TypeDecl};

pub(super) struct InstanceType {
    binary: Arc<ComponentBinary>,
    definition: usize,
    outer: HashMap<u32, u32>,
}

pub(super) struct InstanceInterface {
    funcs: Vec<(String, u32)>,
    types: HashMap<String, u32>,
}

pub(super) fn alias_type(walker: &mut WastWalker, source: u32) -> Result<u32, String> {
    let ty = walker.comp_types.get(source as usize).ok_or("type alias index out of bounds")?;
    let declaration = match ty {
        TypeDecl::Value(_) => TypeDecl::Value(vybe_ast::canon::ValSpec::Ref(source)),
        other => other.clone(),
    };
    let index = u32::try_from(walker.comp_types.len()).map_err(|_| "type alias index overflow")?;
    walker.comp_types.push(declaration);
    walker.comp_type_slots.push(index);
    if let Some(schema) = walker.comp_instance_types.get(&source).cloned() {
        walker.comp_instance_types.insert(index, schema);
    }
    if let Some(interface) = walker.comp_interface_funcs.get(&source).cloned() {
        walker.comp_interface_funcs.insert(index, interface);
    }
    Ok(index)
}

fn type_definitions(definitions: &[Def]) -> Result<Vec<Option<usize>>, String> {
    let mut types = Vec::new();
    for (position, definition) in definitions.iter().enumerate() {
        match definition {
            Def::InterfaceType { .. } | Def::FuncType { .. } | Def::ValueType(_)
            | Def::ResourceType { .. } => types.push(Some(position)),
            Def::ExportType { index, .. }
            | Def::Import { ty: Ref::TypeEq(index), .. } => {
                types.push(*types.get(*index as usize)
                    .ok_or("component type alias index out of bounds")?);
            }
            Def::Import { ty: Ref::TypeSubResource, .. } => types.push(None),
            Def::AliasInstanceExport { kind: vybe_ast::canon::ComponentItemKind::Type, .. } =>
                types.push(None),
            _ => {}
        }
    }
    Ok(types)
}

fn capture_outer(declarations: &[Decl], slots: &[u32], depth: u32,
    captured: &mut HashMap<u32, u32>) -> Result<(), String> {
    for declaration in declarations {
        match declaration {
            Decl::AliasOuterType { count, index } if *count == depth => {
                captured.insert(*index, *slots.get(*index as usize)
                    .ok_or("instance interface: outer type index out of bounds")?);
            }
            Decl::Type(definition) => {
                if let Def::InterfaceType { declarations, .. } = definition.as_ref() {
                    capture_outer(declarations, slots, depth.checked_add(1)
                        .ok_or("instance interface nesting overflow")?, captured)?;
                }
            }
            _ => {}
        }
    }
    Ok(())
}

pub(super) fn register(walker: &mut WastWalker, payload: u32, local: usize) -> Result<(), String> {
    let key = walker.component_binary_scope.ok_or("instance type has no component scope")?;
    let binary = walker.component_interface_binaries.get(&key)
        .ok_or("instance type has no retained standard declarations")?.clone();
    let indices = if let Some(indices) = walker.comp_binary_type_indices.get(&key) {
        indices.clone()
    } else {
        let indices = Arc::new(type_definitions(&binary.definitions)?);
        walker.comp_binary_type_indices.insert(key, indices.clone());
        indices
    };
    let definition = indices.get(local).copied().flatten()
        .ok_or("instance type has no standard definition")?;
    let Some(Def::InterfaceType { is_component: false, declarations }) =
        binary.definitions.get(definition) else {
            return Err("instance type projection does not match its standard type index".into());
        };
    let mut outer = HashMap::new();
    capture_outer(declarations, &walker.comp_type_slots[..local], 1, &mut outer)?;
    walker.comp_instance_types.insert(payload, Arc::new(InstanceType {
        binary, definition, outer,
    }));
    Ok(())
}

fn project(walker: &mut WastWalker, payload: u32) -> Result<(), String> {
    if walker.comp_interface_funcs.contains_key(&payload) { return Ok(()); }
    let source = walker.comp_instance_types.get(&payload)
        .ok_or("import: type is not an instance interface")?.clone();
    let Some(Def::InterfaceType { is_component: false, declarations }) =
        source.binary.definitions.get(source.definition) else {
            return Err("instance interface projection lost its definition".into());
        };
    let mut slots = HashMap::new();
    let mut local_count = 0u32;
    let mut exports = Vec::new();
    let mut exported_types = HashMap::new();
    for declaration in declarations {
        match declaration {
            Decl::Type(definition) => {
                let mapped = |spec| binary_valspec(spec, &slots)
                    .ok_or_else(|| "instance interface: value type reference out of bounds".to_string());
                let ty = match definition.as_ref() {
                    Def::ValueType(spec) => TypeDecl::Value(mapped(spec)?),
                    Def::FuncType { params, result, is_async } => TypeDecl::Func {
                        params: params.iter().map(|(name, spec)|
                            Ok((name.clone(), mapped(spec)?))).collect::<Result<_, String>>()?,
                        result: result.as_ref().map(mapped).transpose()?,
                        is_async: *is_async,
                    },
                    _ => return Err("instance interface: nested interface/resource projection is not implemented".into()),
                };
                validate_component_type_size(&ty, &walker.comp_types)?;
                let index = u32::try_from(walker.comp_types.len())
                    .map_err(|_| "instance interface type index overflow")?;
                walker.comp_types.push(ty);
                slots.insert(local_count, index);
                local_count += 1;
            }
            Decl::AliasOuterType { count: 1, index } => {
                let global = *source.outer.get(index)
                    .ok_or("instance interface: outer type index out of bounds")?;
                slots.insert(local_count, global);
                local_count += 1;
            }
            Decl::Export { name, ty: Ref::Func(index) } => {
                let global = *slots.get(index).ok_or("instance interface function type out of bounds")?;
                if !matches!(walker.comp_types.get(global as usize), Some(TypeDecl::Func { .. })) {
                    return Err("instance interface export does not reference a function type".into());
                }
                exports.push((name.clone(), global));
            }
            Decl::Export { name, ty: Ref::TypeEq(index) } => {
                let global = *slots.get(index).ok_or("instance interface exported type out of bounds")?;
                if !matches!(walker.comp_types.get(global as usize), Some(TypeDecl::Value(_))) {
                    return Err("instance interface: non-value type export requires identity substitution".into());
                }
                exported_types.insert(name.clone(), global);
                slots.insert(local_count, global);
                local_count += 1;
            }
            _ => return Err("instance interface: non-function export or nested lexical alias requires typed projection".into()),
        }
    }
    walker.comp_interface_funcs.insert(payload, Arc::new(InstanceInterface {
        funcs: exports, types: exported_types,
    }));
    Ok(())
}

pub(super) fn import(walker: &mut WastWalker, name: String, ty: u32,
    binder: Option<String>) -> Result<(), String> {
    let local = walker.comp_type_slots.binary_search(&ty)
        .map_err(|_| "type index out of bounds")?;
    match walker.comp_types.get(ty as usize) {
        Some(TypeDecl::Opaque(kind)) if kind == "instance" => {}
        Some(_) => return Err(format!("type index {local} is not an instance type")),
        None => return Err("type index out of bounds".into()),
    }
    project(walker, ty)?;
    let declarations = walker.comp_interface_funcs.get(&ty)
        .ok_or("import: type is not an executable instance interface")?.clone();
    let mut funcs = HashMap::with_capacity(declarations.funcs.len());
    for (name, ty) in &declarations.funcs {
        let index = u32::try_from(walker.comp_func_space.len())
            .map_err(|_| "instance import function index overflow")?;
        walker.comp_func_space.push(CompFunc { canonidx: None, functype: Some(*ty) });
        funcs.insert(name.clone(), index);
    }
    let instance = Arc::new(CompInstance { funcs, types: declarations.types.clone() });
    walker.comp_import_instances.insert(name, instance.clone());
    let index = u32::try_from(walker.comp_instances.len())
        .map_err(|_| "instance import index overflow")?;
    walker.comp_instances.push(instance);
    if let Some(binder) = binder { walker.comp_instance_index.insert(binder, index); }
    Ok(())
}
