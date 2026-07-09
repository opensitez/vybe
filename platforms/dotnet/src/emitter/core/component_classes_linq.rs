use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, MethodBody, MethodDef};

pub(super) fn apply_linq_registrations(exports: &mut [DotnetClassExport]) {
    for export in exports.iter_mut() {
        if is_linq_target(export.interface, &export.class) {
            add_linq_instance_methods(&mut export.class);
        }
    }
}

/// The single `System.Linq.Enumerable` surface. LINQ is a set of extension
/// methods on `IEnumerable<T>`, not methods of any one collection, so it is
/// declared exactly once here. `lookup_instance_method` falls back to this
/// class for ANY enumerable receiver (arrays, `List<T>`, `HashSet<T>`, query
/// results, `yield` generators, and — because the adapters drain through the
/// shared ECMA §7.4 iterator protocol — iterables produced by any Vybe
/// frontend). One declaration, every language, no per-language profile entries.
pub(super) fn enumerable_export() -> DotnetClassExport {
    let mut class = ClassType::new("IEnumerable");
    add_linq_instance_methods(&mut class);
    DotnetClassExport::new("dotnet.System.Collections.Generic", class)
}

fn is_linq_target(interface: &str, class: &ClassType) -> bool {
    (interface == "dotnet.System.Collections.Generic" && class.name == "List")
        || (interface == "dotnet.System.Collections" && class.name == "ArrayList")
}

fn add_linq_instance_methods(class: &mut ClassType) {
    class.methods.push(MethodDef::new(
        "Select",
        1,
        MethodBody::Common("dotnet.linq_select".into()),
    ));
    // `Count()` (no predicate) and `Count(pred)` are distinct overloads.
    // Both must be declared so the (case-insensitive, exact-arity) dotnet
    // resolver matches each — otherwise the zero-arg call falls through to
    // the runtime-collection VM binding, which the exact VM misses for a
    // PascalCase (`Count`) spelling. Resolving here keeps C# and VB on the
    // same case-insensitive resolver path with no VM-level case fold.
    class.methods.push(MethodDef::new(
        "Count",
        0,
        MethodBody::Common("dotnet.linq_count".into()),
    ));
    class.methods.push(MethodDef::new(
        "Count",
        1,
        MethodBody::Common("dotnet.linq_count_pred".into()),
    ));
    class.methods.push(MethodDef::new(
        "Where",
        1,
        MethodBody::Common("dotnet.linq_where".into()),
    ));
    class.methods.push(MethodDef::new(
        "Any",
        0,
        MethodBody::Common("dotnet.linq_any".into()),
    ));
    class.methods.push(MethodDef::new(
        "Contains",
        1,
        MethodBody::Common("dotnet.linq_contains".into()),
    ));
    class.methods.push(MethodDef::new(
        "Reverse",
        0,
        MethodBody::Common("dotnet.linq_reverse".into()),
    ));
    class.methods.push(MethodDef::new(
        "SkipWhile",
        1,
        MethodBody::Common("dotnet.linq_skip_while".into()),
    ));
    class.methods.push(MethodDef::new(
        "TakeWhile",
        1,
        MethodBody::Common("dotnet.linq_take_while".into()),
    ));
    class.methods.push(MethodDef::new(
        "First",
        0,
        MethodBody::Common("dotnet.linq_first".into()),
    ));
    class.methods.push(MethodDef::new(
        "Last",
        0,
        MethodBody::Common("dotnet.linq_last".into()),
    ));
    class.methods.push(MethodDef::new(
        "Skip",
        1,
        MethodBody::Common("dotnet.linq_skip".into()),
    ));
    class.methods.push(MethodDef::new(
        "Take",
        1,
        MethodBody::Common("dotnet.linq_take".into()),
    ));
    class.methods.push(MethodDef::new(
        "Average",
        0,
        MethodBody::Common("dotnet.linq_average".into()),
    ));
    class.methods.push(MethodDef::new(
        "FirstOrDefault",
        0,
        MethodBody::Common("dotnet.linq_first_or_default".into()),
    ));
    class.methods.push(MethodDef::new(
        "Distinct",
        0,
        MethodBody::Common("dotnet.linq_distinct".into()),
    ));
    class.methods.push(MethodDef::new(
        "Aggregate",
        2,
        MethodBody::Common("dotnet.linq_aggregate".into()),
    ));
    class.methods.push(MethodDef::new(
        "OrderBy",
        1,
        MethodBody::Common("dotnet.linq_order_by".into()),
    ));
    class.methods.push(MethodDef::new(
        "OrderByDescending",
        1,
        MethodBody::Common("dotnet.linq_order_by_descending".into()),
    ));
    class.methods.push(MethodDef::new(
        "DistinctBy",
        1,
        MethodBody::Common("dotnet.linq_distinct_by".into()),
    ));
    class.methods.push(MethodDef::new(
        "GroupBy",
        1,
        MethodBody::Common("dotnet.linq_group_by".into()),
    ));
    class.methods.push(MethodDef::new(
        "SelectMany",
        1,
        MethodBody::Common("dotnet.linq_select_many".into()),
    ));
    class.methods.push(MethodDef::new(
        "ToDictionary",
        2,
        MethodBody::Common("dotnet.linq_to_dictionary".into()),
    ));
    class.methods.push(MethodDef::new(
        "Zip",
        2,
        MethodBody::Common("dotnet.linq_zip".into()),
    ));
    class.methods.push(MethodDef::new(
        "ToList",
        0,
        MethodBody::Common("dotnet.linq_identity".into()),
    ));
    class.methods.push(MethodDef::new(
        "ToArray",
        0,
        MethodBody::Common("dotnet.linq_identity".into()),
    ));
    class.methods.push(MethodDef::new(
        "Aggregate",
        1,
        MethodBody::Common("dotnet.linq_aggregate_no_seed".into()),
    ));
    class.methods.push(MethodDef::new(
        "Aggregate",
        2,
        MethodBody::Common("dotnet.linq_aggregate".into()),
    ));
    class.methods.push(MethodDef::new(
        "ElementAt",
        1,
        MethodBody::Common("dotnet.linq_element_at".into()),
    ));
    class.methods.push(MethodDef::new(
        "ElementAtOrDefault",
        1,
        MethodBody::Common("dotnet.linq_element_at_or_default".into()),
    ));
    class.methods.push(MethodDef::new(
        "Single",
        0,
        MethodBody::Common("dotnet.linq_single".into()),
    ));
    class.methods.push(MethodDef::new(
        "SingleOrDefault",
        0,
        MethodBody::Common("dotnet.linq_single_or_default".into()),
    ));
    class.methods.push(MethodDef::new(
        "MaxBy",
        1,
        MethodBody::Common("dotnet.linq_max_by".into()),
    ));
    class.methods.push(MethodDef::new(
        "MinBy",
        1,
        MethodBody::Common("dotnet.linq_min_by".into()),
    ));
    class.methods.push(MethodDef::new(
        "Append",
        1,
        MethodBody::Common("dotnet.linq_append".into()),
    ));
    class.methods.push(MethodDef::new(
        "Prepend",
        1,
        MethodBody::Common("dotnet.linq_prepend".into()),
    ));
    class.methods.push(MethodDef::new(
        "SkipLast",
        1,
        MethodBody::Common("dotnet.linq_skip_last".into()),
    ));
    class.methods.push(MethodDef::new(
        "TakeLast",
        1,
        MethodBody::Common("dotnet.linq_take_last".into()),
    ));
    class.methods.push(MethodDef::new(
        "DefaultIfEmpty",
        0,
        MethodBody::Common("dotnet.linq_default_if_empty".into()),
    ));
    // Aggregation. `Sum` drains any receiver (array, `List<T>`, generator,
    // cross-language iterable) through the shared iterator protocol before
    // summing; `Min`/`Max` operate on the materialized sequence.
    class.methods.push(MethodDef::new(
        "Sum",
        0,
        MethodBody::Common("dotnet.linq_sum".into()),
    ));
    class.methods.push(MethodDef::new(
        "Min",
        0,
        MethodBody::Common("collections.min".into()),
    ));
    class.methods.push(MethodDef::new(
        "Max",
        0,
        MethodBody::Common("collections.max".into()),
    ));
    class.methods.push(MethodDef::new(
        "SequenceEqual",
        1,
        MethodBody::Common("dotnet.linq_sequence_equal".into()),
    ));
}
