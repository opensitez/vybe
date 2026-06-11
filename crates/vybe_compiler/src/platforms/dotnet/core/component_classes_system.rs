use super::super::super::class_exports::DotnetClassExport;
use super::super::super::classes::DotnetClass;
use vybe_bytecode::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef, PropertyDef,
};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::with_wrapper(
            "dotnet.System",
            ClassType::new("ValueType").with_parent("Object"),
            DotnetClass {
                name: "ValueType",
                parent: Some("Object"),
                properties: &[],
                methods: &[],
                ctor_arity: 0,
                widget_host_fn: None,
                widget_host_module: "vybe:gui",
            },
        ),
        DotnetClassExport::with_wrapper(
            "dotnet.System",
            ClassType::new("Enum").with_parent("ValueType"),
            DotnetClass {
                name: "Enum",
                parent: Some("ValueType"),
                properties: &[],
                methods: &[],
                ctor_arity: 0,
                widget_host_fn: None,
                widget_host_module: "vybe:gui",
            },
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("DateTime")
                .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.datetime_new"))
                .with_method(MethodDef::static_method(
                    "Now",
                    0,
                    MethodBody::Common("dotnet.datetime_now".into()),
                ))
                .with_method(MethodDef::static_method(
                    "UtcNow",
                    0,
                    MethodBody::Common("dotnet.datetime_now".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Today",
                    0,
                    MethodBody::Common("dotnet.datetime_today".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Parse",
                    1,
                    MethodBody::Common("dotnet.datetime_parse".into()),
                ))
                .with_method(MethodDef::static_method(
                    "DaysInMonth",
                    2,
                    MethodBody::Common("dotnet.datetime_days_in_month".into()),
                ))
                .with_method(MethodDef::static_method(
                    "IsLeapYear",
                    1,
                    MethodBody::Common("dotnet.datetime_is_leap_year".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Compare",
                    2,
                    MethodBody::Common("dotnet.datetime_compare".into()),
                ))
                .with_method(MethodDef::new(
                    "AddDays",
                    1,
                    MethodBody::Common("dotnet.datetime_add_days".into()),
                ))
                .with_method(MethodDef::new(
                    "AddHours",
                    1,
                    MethodBody::Common("dotnet.datetime_add_hours".into()),
                ))
                .with_method(MethodDef::new(
                    "AddMonths",
                    1,
                    MethodBody::Common("dotnet.datetime_add_months".into()),
                ))
                .with_method(MethodDef::new(
                    "ToShortDateString",
                    0,
                    MethodBody::Common("dotnet.datetime_to_short_date_string".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Random")
                .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.random_new"))
                .with_method(MethodDef::new(
                    "Next",
                    0,
                    MethodBody::Common("dotnet.random_next".into()),
                ))
                .with_method(MethodDef::new(
                    "Next",
                    1,
                    MethodBody::Common("dotnet.random_next".into()),
                ))
                .with_method(MethodDef::new(
                    "Next",
                    2,
                    MethodBody::Common("dotnet.random_next".into()),
                ))
                .with_method(MethodDef::new(
                    "NextDouble",
                    0,
                    MethodBody::Common("dotnet.random_next_double".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("TimeSpan")
                .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.timespan_new"))
                .with_method(MethodDef::static_method(
                    "FromDays",
                    1,
                    MethodBody::Common("dotnet.timespan_from_days".into()),
                ))
                .with_method(MethodDef::static_method(
                    "FromHours",
                    1,
                    MethodBody::Common("dotnet.timespan_from_hours".into()),
                ))
                .with_method(MethodDef::static_method(
                    "FromMinutes",
                    1,
                    MethodBody::Common("dotnet.timespan_from_minutes".into()),
                ))
                .with_method(MethodDef::static_method(
                    "FromSeconds",
                    1,
                    MethodBody::Common("dotnet.timespan_from_seconds".into()),
                ))
                .with_method(MethodDef::static_method(
                    "FromMilliseconds",
                    1,
                    MethodBody::Common("dotnet.timespan_from_milliseconds".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Zero",
                    0,
                    MethodBody::Common("dotnet.timespan_zero".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Compare",
                    2,
                    MethodBody::Common("dotnet.timespan_compare".into()),
                ))
                .with_method(MethodDef::new(
                    "Negate",
                    0,
                    MethodBody::Common("dotnet.timespan_negate".into()),
                ))
                .with_method(MethodDef::new(
                    "Duration",
                    0,
                    MethodBody::Common("dotnet.timespan_duration".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Array")
                .with_method(MethodDef::static_method(
                    "Clear",
                    3,
                    MethodBody::Common("dotnet.array_clear".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Copy",
                    3,
                    MethodBody::Common("dotnet.array_copy".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Resize",
                    2,
                    MethodBody::Common("dotnet.array_resize".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Sort",
                    1,
                    MethodBody::Common("dotnet.array_sort".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Reverse",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "reverse")),
                ))
                .with_method(MethodDef::static_method(
                    "IndexOf",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "indexOf")),
                ))
                .with_method(MethodDef::static_method(
                    "LastIndexOf",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "lastIndexOf")),
                ))
                .with_method(MethodDef::static_method(
                    "Empty",
                    0,
                    MethodBody::Common("collections.new".into()),
                ))
                .with_method(MethodDef::static_method(
                    "BinarySearch",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "indexOf")),
                ))
                .with_method(MethodDef::static_method(
                    "ConvertAll",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "map")),
                ))
                .with_method(MethodDef::static_method(
                    "CreateInstance",
                    1,
                    MethodBody::Common("collections.new".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Exists",
                    2,
                    MethodBody::Common("dotnet.array_exists".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Find",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "find")),
                ))
                .with_method(MethodDef::static_method(
                    "FindIndex",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "findIndex")),
                ))
                .with_method(MethodDef::static_method(
                    "FindLastIndex",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "findIndex")),
                ))
                .with_method(MethodDef::static_method(
                    "ForEach",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "forEach")),
                ))
                .with_method(MethodDef::static_method(
                    "TrueForAll",
                    2,
                    MethodBody::Common("dotnet.array_true_for_all".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Console")
                .with_method(MethodDef::static_method(
                    "WriteLine",
                    1,
                    MethodBody::Common("dotnet.console_writeline".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Write",
                    1,
                    MethodBody::Common("dotnet.console_writeline".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadLine",
                    0,
                    MethodBody::Common("dotnet.console_readline".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Error",
                    1,
                    MethodBody::Common("dotnet.console_error".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Print",
                    1,
                    MethodBody::Common("dotnet.console_writeline".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Assert",
                    1,
                    MethodBody::Common("dotnet.console_writeline".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Delegate")
                .with_method(MethodDef::static_method(
                    "Combine",
                    2,
                    MethodBody::Common("delegates.combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Remove",
                    2,
                    MethodBody::Common("delegates.remove".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Convert")
                .with_method(MethodDef::static_method(
                    "ToInt32",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:number", "parseInt")),
                ))
                .with_method(MethodDef::static_method(
                    "ToDouble",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:number", "Number")),
                ))
                .with_method(MethodDef::static_method(
                    "ToString",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:string", "String")),
                ))
                .with_method(MethodDef::static_method(
                    "ToBoolean",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:boolean", "Boolean")),
                ))
                .with_method(MethodDef::static_method(
                    "ToDateTime",
                    1,
                    MethodBody::Common("dotnet.datetime_parse".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("String").with_method(MethodDef::static_method(
                "Format",
                2,
                MethodBody::Common("dotnet.string_format".into()),
            )),
        ),
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Environment")
                .with_property(
                    PropertyDef::new("CurrentDirectory")
                        .with_getter(HostTarget::new("node:process", "cwd")),
                )
                .with_property(
                    PropertyDef::new("NewLine").with_getter(HostTarget::new("node:os", "EOL")),
                )
                .with_property(
                    PropertyDef::new("MachineName")
                        .with_getter(HostTarget::new("node:os", "hostname")),
                )
                .with_property(
                    PropertyDef::new("OSVersion")
                        .with_getter(HostTarget::new("node:os", "version")),
                )
                .with_method(MethodDef::static_method(
                    "UserName",
                    0,
                    MethodBody::Common("dotnet.environment_username".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ProcessorCount",
                    0,
                    MethodBody::Common("dotnet.environment_processor_count".into()),
                ))
                .with_method(MethodDef::static_method(
                    "TickCount",
                    0,
                    MethodBody::Common("dotnet.environment_tick_count".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetEnvironmentVariable",
                    1,
                    MethodBody::Common("dotnet.environment_get".into()),
                ))
                .with_method(MethodDef::static_method(
                    "SetEnvironmentVariable",
                    2,
                    MethodBody::Common("dotnet.environment_set".into()),
                )),
        ),
    ]
}
