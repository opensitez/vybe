use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, MethodBody, MethodDef};

pub(super) fn apply_linq_registrations(exports: &mut [DotnetClassExport]) {
    for export in exports.iter_mut() {
        if is_linq_target(export.interface, &export.class) {
            add_linq_instance_methods(&mut export.class);
        }
    }
}

fn is_linq_target(interface: &str, class: &ClassType) -> bool {
    (interface == "dotnet.System.Collections.Generic" && class.name == "List")
        || (interface == "dotnet.System.Collections" && class.name == "ArrayList")
}

fn add_linq_instance_methods(class: &mut ClassType) {
    class.methods.push(MethodDef::new("Select", 1, MethodBody::Common("dotnet.linq_select".into())));
    class.methods.push(MethodDef::new("Count", 1, MethodBody::Common("dotnet.linq_count_pred".into())));
    class.methods.push(MethodDef::new("First", 0, MethodBody::Common("dotnet.linq_first".into())));
    class.methods.push(MethodDef::new("Last", 0, MethodBody::Common("dotnet.linq_last".into())));
    class.methods.push(MethodDef::new("Skip", 1, MethodBody::Common("dotnet.linq_skip".into())));
    class.methods.push(MethodDef::new("Take", 1, MethodBody::Common("dotnet.linq_take".into())));
    class.methods.push(MethodDef::new("Average", 0, MethodBody::Common("dotnet.linq_average".into())));
    class.methods.push(MethodDef::new("FirstOrDefault", 0, MethodBody::Common("dotnet.linq_first_or_default".into())));
    class.methods.push(MethodDef::new("Distinct", 0, MethodBody::Common("dotnet.linq_distinct".into())));
    class.methods.push(MethodDef::new("Aggregate", 2, MethodBody::Common("dotnet.linq_aggregate".into())));
    class.methods.push(MethodDef::new("OrderByDescending", 1, MethodBody::Common("dotnet.linq_order_by_descending".into())));
    class.methods.push(MethodDef::new("GroupBy", 1, MethodBody::Common("dotnet.linq_group_by".into())));
    class.methods.push(MethodDef::new("SelectMany", 1, MethodBody::Common("dotnet.linq_select_many".into())));
    class.methods.push(MethodDef::new("ToDictionary", 2, MethodBody::Common("dotnet.linq_to_dictionary".into())));
    class.methods.push(MethodDef::new("Zip", 2, MethodBody::Common("dotnet.linq_zip".into())));
    class.methods.push(MethodDef::new("ToList", 0, MethodBody::Common("dotnet.linq_identity".into())));
}
