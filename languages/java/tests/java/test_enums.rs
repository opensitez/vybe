use crate::helpers::run_in_main;

#[test]
fn enum_constants_have_declared_order() {
    let types = r#"
        enum Day { MON, TUE, WED }
    "#;
    let out = run_in_main(
        "System.out.println(Day.MON); System.out.println(Day.WED);",
        types,
    );
    assert_eq!(out, vec!["MON", "WED"]);
}

#[test]
fn enum_ordinal_reflects_declaration_index() {
    let types = r#"
        enum Size { SMALL, MEDIUM, LARGE }
    "#;
    let out = run_in_main("System.out.println(Size.MEDIUM.ordinal());", types);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_switch_selects_matching_constant() {
    let types = r#"
        enum Color { RED, GREEN, BLUE }
    "#;
    let out = run_in_main(
        "Color c = Color.GREEN; switch (c) { case RED: System.out.println(\"r\"); break; case GREEN: System.out.println(\"g\"); break; default: System.out.println(\"b\"); }",
        types,
    );
    assert_eq!(out, vec!["g"]);
}

#[test]
fn enum_with_fields_exposes_custom_payload() {
    let types = r#"
        enum Planet { EARTH(3), MARS(2);
            final int moons;
            Planet(int m) { moons = m; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Planet.EARTH.moons); System.out.println(Planet.MARS.moons);",
        types,
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn enum_value_of_parses_name() {
    let types = r#"
        enum Status { ON, OFF }
    "#;
    let out = run_in_main("System.out.println(Status.valueOf(\"ON\"));", types);
    assert_eq!(out, vec!["ON"]);
}
