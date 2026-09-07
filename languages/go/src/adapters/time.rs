use std::collections::BTreeMap;

use vybe_ast::{Argument, ExprKind, Expression, Literal};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};
use vybe_runtime::Value;

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("time.Date", "go.time_date"),
        ("time.Unix", "go.time_unix"),
        ("time.Now", "go.time_now"),
        ("time.UnixMilli", "go.time_unix_milli"),
        ("time.UnixMicro", "go.time_unix_micro"),
        ("time.FixedZone", "go.time_fixed_zone"),
        ("time.LoadLocation", "go.time_load_location"),
        ("time.Parse", "go.time_parse"),
        ("time.ParseInLocation", "go.time_parse_in_location"),
        ("time.ParseDuration", "go.time_parse_duration"),
        ("time.Since", "go.time_since"),
        ("time.Until", "go.time_until"),
        ("time.Sleep", "go.time_sleep"),
        ("time.After", "go.time_after"),
        ("time.UTC", "go.time_utc"),
        ("time.Local", "go.time_local"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    for (name, value) in [
        ("time.Sunday", "Sunday"),
        ("time.Monday", "Monday"),
        ("time.Tuesday", "Tuesday"),
        ("time.Wednesday", "Wednesday"),
        ("time.Thursday", "Thursday"),
        ("time.Friday", "Friday"),
        ("time.Saturday", "Saturday"),
        ("time.January", "January"),
        ("time.February", "February"),
        ("time.March", "March"),
        ("time.April", "April"),
        ("time.May", "May"),
        ("time.June", "June"),
        ("time.July", "July"),
        ("time.August", "August"),
        ("time.September", "September"),
        ("time.October", "October"),
        ("time.November", "November"),
        ("time.December", "December"),
        ("time.RFC3339", "2006-01-02T15:04:05Z07:00"),
        ("time.RFC822", "02 Jan 06 15:04 MST"),
        ("time.Kitchen", "3:04PM"),
        ("time.UnixDate", "Mon Jan _2 15:04:05 MST 2006"),
        ("time.Stamp", "Jan _2 15:04:05"),
        ("time.StampMicro", "Jan _2 15:04:05.000000"),
    ] {
        insert_path(
            root,
            name,
            NamespaceNode::Const(Value::String(value.into())),
        );
    }

    for (name, value) in [
        ("time.Nanosecond", 1.0),
        ("time.Microsecond", 1000.0),
        ("time.Millisecond", 1000000.0),
        ("time.Second", 1000000000.0),
        ("time.Minute", 60000000000.0),
        ("time.Hour", 3600000000000.0),
    ] {
        insert_path(root, name, NamespaceNode::Const(Value::F64(value)));
    }

    let mut time_methods = Subtree::new();
    for method in [
        "Format",
        "Year",
        "AddDate",
        "Add",
        "Sub",
        "Month",
        "MonthInt",
        "Day",
        "Hour",
        "Minute",
        "Second",
        "Nanosecond",
        "Unix",
        "UnixNano",
        "UnixMilli",
        "UnixMicro",
        "Weekday",
        "YearDay",
        "Zone",
        "Before",
        "After",
        "Equal",
        "Truncate",
        "Round",
        "UTC",
        "In",
        "Location",
        "IsZero",
    ] {
        time_methods.insert(
            method.to_string(),
            NamespaceNode::CommonEmit(format!("go.time.Time.{method}")),
        );
    }
    insert_path(
        root,
        "time.Time",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: time_methods,
            member_returns: [
                ("AddDate".to_string(), "time.Time".to_string()),
                ("Add".to_string(), "time.Time".to_string()),
                ("Round".to_string(), "time.Time".to_string()),
                ("Truncate".to_string(), "time.Time".to_string()),
                ("UTC".to_string(), "time.Time".to_string()),
                ("In".to_string(), "time.Time".to_string()),
                ("Location".to_string(), "time.Location".to_string()),
            ]
            .into_iter()
            .collect(),
        },
    );

    let mut loc_methods = Subtree::new();
    loc_methods.insert(
        "String".to_string(),
        NamespaceNode::CommonEmit("go.time.Location.String".to_string()),
    );
    insert_path(
        root,
        "time.Location",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: loc_methods,
            member_returns: BTreeMap::new(),
        },
    );

    let mut duration_methods = Subtree::new();
    for method in [
        "String",
        "Round",
        "Minutes",
        "Seconds",
        "Hours",
        "Nanoseconds",
        "Milliseconds",
        "Microseconds",
    ] {
        if let Some(emit) = duration_method_emit(method) {
            duration_methods.insert(
                method.to_string(),
                NamespaceNode::CommonEmit(emit.to_string()),
            );
        }
    }
    insert_path(
        root,
        "time.Duration",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: duration_methods,
            member_returns: [("Round".to_string(), "time.Duration".to_string())]
                .into_iter()
                .collect(),
        },
    );
}

fn duration_method_emit(method: &str) -> Option<&'static str> {
    match method {
        "String" => Some("go.time.Duration.String"),
        "Round" => Some("go.time.Duration.Round"),
        "Minutes" => Some("go.dur_minutes"),
        "Seconds" => Some("go.dur_seconds"),
        "Hours" => Some("go.dur_hours"),
        "Nanoseconds" => Some("go.dur_nanoseconds"),
        "Milliseconds" => Some("go.dur_milliseconds"),
        "Microseconds" => Some("go.dur_microseconds"),
        _ => None,
    }
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "time.Time" => Some("__goTime"),
        "time.Location" => Some("__goLoc"),
        _ => None,
    }
}

pub(crate) fn rewrite_member(field: &str) -> Option<Expression> {
    match field {
        "UTC" => Some(go_builtin_call("__go_time_utc", Vec::new())),
        "Local" => Some(go_builtin_call("__go_time_local", Vec::new())),
        "Sunday" => Some(Expression::string("Sunday")),
        "Monday" => Some(Expression::string("Monday")),
        "Tuesday" => Some(Expression::string("Tuesday")),
        "Wednesday" => Some(Expression::string("Wednesday")),
        "Thursday" => Some(Expression::string("Thursday")),
        "Friday" => Some(Expression::string("Friday")),
        "Saturday" => Some(Expression::string("Saturday")),
        "January" => Some(Expression::string("January")),
        "February" => Some(Expression::string("February")),
        "March" => Some(Expression::string("March")),
        "April" => Some(Expression::string("April")),
        "May" => Some(Expression::string("May")),
        "June" => Some(Expression::string("June")),
        "July" => Some(Expression::string("July")),
        "August" => Some(Expression::string("August")),
        "September" => Some(Expression::string("September")),
        "October" => Some(Expression::string("October")),
        "November" => Some(Expression::string("November")),
        "December" => Some(Expression::string("December")),
        "RFC3339" => Some(Expression::string("2006-01-02T15:04:05Z07:00")),
        "RFC822" => Some(Expression::string("02 Jan 06 15:04 MST")),
        "Kitchen" => Some(Expression::string("3:04PM")),
        "UnixDate" => Some(Expression::string("Mon Jan _2 15:04:05 MST 2006")),
        "Stamp" => Some(Expression::string("Jan _2 15:04:05")),
        "StampMicro" => Some(Expression::string("Jan _2 15:04:05.000000")),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if call_name == "time.Date" {
        let mut values = args.iter().map(|arg| arg.value.clone()).collect::<Vec<_>>();
        if values.len() >= 2
            && let Some(month) = named_value_to_int(&values[1])
        {
            values[1] = Expression::int(month);
        }
        return Some(go_builtin_call("__go_time_date", values));
    }
    let mapped = match call_name {
        "time.Unix" => "__go_time_unix",
        "time.Now" => "__go_time_now",
        "time.UnixMilli" => "__go_time_unix_milli",
        "time.UnixMicro" => "__go_time_unix_micro",
        "time.FixedZone" => "__go_time_fixed_zone",
        "time.LoadLocation" => "__go_time_load_location",
        "time.Parse" => "__go_time_parse",
        "time.ParseInLocation" => "__go_time_parse_in_location",
        "time.ParseDuration" => "__go_time_parse_duration",
        "time.Since" => "__go_time_since",
        "time.Until" => "__go_time_until",
        "time.Sleep" => "__go_time_sleep",
        "time.After" => "__go_time_after",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|arg| arg.value.clone()).collect(),
    ))
}

fn named_value_to_int(expr: &Expression) -> Option<i64> {
    let ExprKind::Lit(Literal::Str(name)) = &expr.kind else {
        return None;
    };
    match name.as_str() {
        "Sunday" => Some(0),
        "Monday" => Some(1),
        "Tuesday" => Some(2),
        "Wednesday" => Some(3),
        "Thursday" => Some(4),
        "Friday" => Some(5),
        "Saturday" => Some(6),
        "January" => Some(1),
        "February" => Some(2),
        "March" => Some(3),
        "April" => Some(4),
        "May" => Some(5),
        "June" => Some(6),
        "July" => Some(7),
        "August" => Some(8),
        "September" => Some(9),
        "October" => Some(10),
        "November" => Some(11),
        "December" => Some(12),
        _ => None,
    }
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.insert(leaf.to_string(), node);
}
