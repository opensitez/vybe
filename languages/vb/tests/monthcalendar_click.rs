//! Actual VB project, real browser hit testing and the registered guest callbacks.
use vybe_platform_web::{
    engine::{self, DocumentId, DomOp, DomValue, NodeId},
    html,
};
use vybe_runtime::{VM, Value};
fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .ancestors()
        .find(|p| {
            p.join("examples/vb/allcontrols/allcontrols.vbproj")
                .is_file()
        })
        .expect("repository-backed manifest")
        .to_path_buf()
}
fn node(document: DocumentId, selector: &str) -> NodeId {
    match engine::apply(document, DomOp::QuerySelector(selector.into())) {
        DomValue::Node(n) => n,
        x => panic!("real DOM selector {selector}: {x:?}"),
    }
}
fn child(document: DocumentId, parent: NodeId, selector: &str) -> NodeId {
    match engine::apply(
        document,
        DomOp::QuerySelectorWithin {
            root: parent,
            selector: selector.into(),
        },
    ) {
        DomValue::Node(n) => n,
        x => panic!("real calendar selector {selector}: {x:?}"),
    }
}
fn attr(document: DocumentId, node: NodeId, name: &str) -> String {
    match engine::apply(document, DomOp::GetAttribute(node, name.into())) {
        DomValue::Text(s) => s,
        x => panic!("attribute {name}: {x:?}"),
    }
}
fn click(vm: &mut VM, document: DocumentId, target: NodeId) {
    let DomValue::Rect {
        x,
        y,
        width,
        height,
    } = engine::apply(document, DomOp::BoundingClientRect(target))
    else {
        panic!("real target has no box")
    };
    assert!(width > 0. && height > 0., "real target must be laid out");
    for kind in ["mousedown", "mouseup"] {
        engine::apply(
            document,
            DomOp::DispatchPointer {
                kind: kind.into(),
                client_x: (x + width / 2.) as f32,
                client_y: (y + height / 2.) as f32,
                button: 0,
            },
        );
    }
    let dispatches = html::pending_dispatches(document);
    assert!(
        !dispatches.is_empty(),
        "actual engine pointer must deliver registered listeners"
    );
    let mut delivered = false;
    for (callback, event) in dispatches {
        if let Value::Object(o) = &event {
            let o = o.lock().unwrap();
            if matches!(o.properties.get("type"),Some(Value::String(s))if s.as_ref()=="click") {
                delivered = true;
            }
        }
        vm.invoke(&callback, &[event])
            .expect("original registered guest callback succeeds");
        assert!(
            vm.last_exception.is_none(),
            "click must not leave guest exception"
        );
    }
    assert!(delivered, "click event reached original listeners");
}
fn setting_number(control: &vybe_runtime::value::Object, key: &str) -> f64 {
    match control.properties.get(key) {
        Some(Value::I32(v)) => *v as f64,
        Some(Value::F64(v)) => *v,
        x => panic!("actual numeric setting {key}: {x:?}"),
    }
}
fn style(document: DocumentId, node: NodeId, key: &str) -> String {
    match engine::apply(document, DomOp::GetStyleProperty(node, key.into())) {
        DomValue::Text(s) => s,
        x => panic!("style {key}: {x:?}"),
    }
}
fn period(document: DocumentId, calendar: NodeId) -> (i32, i32) {
    (
        attr(document, calendar, "data-year").parse().unwrap(),
        attr(document, calendar, "data-month").parse().unwrap(),
    )
}
fn shifted((y, m): (i32, i32), step: i32) -> (i32, i32) {
    let index = y * 12 + m + step;
    (index.div_euclid(12), index.rem_euclid(12))
}
#[test]
fn actual_vb_calendar_day_navigation_and_live_settings() {
    assert!(
        vybe_platform_web::engine_select::Engine::WebCore.is_available(),
        "regression requires actual webcore engine, not substitute hosts"
    );
    vybe_platform_web::engine_select::choose(vybe_platform_web::engine_select::Engine::WebCore);
    vybe_language_vb::register();
    vybe_platform_dotnet::register();
    let repo = root();
    let project = repo.join("tests/vb/vb_winforms_monthcalendar_click/CalendarClick.vbproj");
    let bundle = vybe_compiler::projects::load(&project).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let prepared = bundle.prepared_module().unwrap();
    let compiled = bundle
        .compile_prepared_module(&prepared, &vm.modules)
        .unwrap();
    assert!(
        compiled
            .chunks
            .iter()
            .any(|c| c.name == "__dotnet_monthcalendar_click"),
        "real calendar adapter emitted"
    );
    vm.run(compiled.chunks).unwrap();
    assert!(vm.last_exception.is_none());
    let document = html::active_document();
    let calendar = node(document, "#calendarProbe");
    let settings = node(document, "#calendarSettings");
    let Value::Object(registry) = vm
        .global("__dotnet_monthcalendar_controls")
        .expect("retained real control registry")
    else {
        panic!("registry must be Object")
    };
    let control = registry
        .lock()
        .unwrap()
        .properties
        .get(&calendar.to_string())
        .cloned()
        .expect("actual node retains its real control");
    let Value::Object(control) = control else {
        panic!("real control object")
    };
    assert_eq!(
        setting_number(&control.lock().unwrap(), "scrollchange"),
        2.,
        "VB initial settings on original control"
    );
    let initial = period(document, calendar);
    let day = child(document, calendar, "[data-day='17']");
    click(&mut vm, document, day);
    assert_eq!(attr(document, calendar, "data-selected"), "17");
    assert_eq!(
        attr(document, calendar, "data-selected-year"),
        initial.0.to_string()
    );
    assert_eq!(
        attr(document, calendar, "data-selected-month"),
        initial.1.to_string()
    );
    let title = child(document, calendar, ".vybe-cal-title");
    click(&mut vm, document, title);
    assert_eq!(
        period(document, calendar),
        initial,
        "non-action title click preserves visible month"
    );
    assert_eq!(
        attr(document, calendar, "data-selected"),
        "17",
        "missing data-day is false, not Undefined at IF"
    );
    let prev = child(document, calendar, "[data-action='prev']");
    click(&mut vm, document, prev);
    assert_eq!(period(document, calendar), shifted(initial, -2));
    let next = child(document, calendar, "[data-action='next']");
    click(&mut vm, document, next);
    assert_eq!(period(document, calendar), initial);
    click(&mut vm, document, settings);
    {
        let c = control.lock().unwrap();
        assert_eq!(setting_number(&c, "scrollchange"), 3.);
        for key in ["showtoday", "showtodaycircle", "showweeknumbers"] {
            assert_eq!(
                c.properties.get(key),
                Some(&Value::Bool(true)),
                "actual VB handler updates {key}"
            );
        }
        assert_eq!(
            c.properties.get("firstdayofweek"),
            Some(&Value::String("Monday".into()))
        );
    }
    let grid = child(document, calendar, ".vybe-cal-days");
    assert_eq!(
        style(document, grid, "grid-template-columns").replace(' ', ""),
        "repeat(8,1fr)",
        "updated Bool enables real week-number layout"
    );
    let footer = child(document, calendar, ".vybe-cal-today");
    assert_eq!(
        style(document, footer, "display"),
        "block",
        "updated ShowToday renders footer"
    );
    let monday = child(document, calendar, ".vybe-cal-weekdays span:nth-child(2)");
    assert!(
        matches!(engine::apply(document,DomOp::TextContent(monday)),DomValue::Text(ref s)if s=="Mo"),
        "Monday setting renders weekday labels"
    );
    let next = child(document, calendar, "[data-action='next']");
    click(&mut vm, document, next);
    assert_eq!(
        period(document, calendar),
        shifted(initial, 3),
        "callback rereads updated original control rather than a stale event wrapper"
    );
    let prev = child(document, calendar, "[data-action='prev']");
    click(&mut vm, document, prev);
    assert_eq!(period(document, calendar), initial);
    assert_eq!(
        attr(document, calendar, "data-selected"),
        "17",
        "navigation preserves selected day"
    );
}

#[test]
fn actual_allcontrols_calendar_click_reaches_real_retained_control() {
    assert!(vybe_platform_web::engine_select::Engine::WebCore.is_available());
    vybe_platform_web::engine_select::choose(vybe_platform_web::engine_select::Engine::WebCore);
    vybe_language_vb::register();
    vybe_platform_dotnet::register();
    let project = root().join("examples/vb/allcontrols/allcontrols.vbproj");
    let bundle = vybe_compiler::projects::load(&project).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let prepared = bundle.prepared_module().unwrap();
    let compiled = bundle
        .compile_prepared_module(&prepared, &vm.modules)
        .unwrap();
    assert!(
        compiled
            .chunks
            .iter()
            .any(|c| c.name == "__dotnet_monthcalendar_click")
    );
    vm.run(compiled.chunks).unwrap();
    assert!(vm.last_exception.is_none());
    let document = html::active_document();
    // Allcontrols places mc1 at y760. Give the real headless browser the
    // corresponding visible client area; no DOM nodes/events are substituted.
    vybe_platform_web::with_browser(document, |doc| doc.set_viewport(800., 1100.))
        .expect("real webcore document");
    let calendar = node(document, "#mc1");
    let initial = period(document, calendar);
    let Value::Object(registry) = vm.global("__dotnet_monthcalendar_controls").unwrap() else {
        panic!("real retained calendar registry")
    };
    assert!(matches!(
        registry
            .lock()
            .unwrap()
            .properties
            .get(&calendar.to_string()),
        Some(Value::Object(_))
    ));
    let selected = attr(document, calendar, "data-selected");
    let day = if selected == "17" { "18" } else { "17" };
    let target = child(document, calendar, &format!("[data-day='{day}']"));
    click(&mut vm, document, target);
    assert_eq!(
        attr(document, calendar, "data-selected"),
        day,
        "actual allcontrols day changes"
    );
    assert_eq!(
        attr(document, calendar, "data-selected-year"),
        initial.0.to_string()
    );
    assert_eq!(
        attr(document, calendar, "data-selected-month"),
        initial.1.to_string()
    );
}
