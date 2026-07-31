use crate::helpers::run_prints;

#[test]
fn test_interface_default_conflict_is_resolved_with_explicit_super_calls() {
    let out = run_prints(r#"
        interface First {
            fun origin(): String = "first"
        }

        interface Second {
            fun origin(): String = "second"
        }

        class Composite : First, Second {
            override fun origin(): String = super<First>.origin() + "/" + super<Second>.origin()
        }

        fun main() {
            println(Composite().origin())
        }
    "#);
    assert_eq!(out, &["first/second"]);
}

#[test]
fn test_interface_property_conflict_is_resolved_by_override() {
    let out = run_prints(r#"
        interface Marker {
            val label: String
                get() = "marker"
        }

        interface Debug {
            val label: String
                get() = "debug"
        }

        class Item : Marker, Debug {
            override val label: String = "item"
        }

        fun main() {
            println(Item().label)
        }
    "#);
    assert_eq!(out, &["item"]);
}
