use crate::helpers::run_prints;

#[test]
fn test_property_setter_normalizes_invalid_values() {
    let out = run_prints(r#"
        class Metric {
            private var _count: Int = 0
            var count: Int
                get() = _count
                set(value) {
                    _count = if (value < 0) 0 else value
                }
        }

        fun main() {
            val m = Metric()
            m.count = -5
            println(m.count)
            m.count = 8
            println(m.count)
        }
    "#);
    assert_eq!(out, &["0", "8"]);
}

#[test]
fn test_property_getter_derived_field() {
    let out = run_prints(r#"
        class Timer {
            private var total = 0
            var seconds: Int
                get() = total
                set(value) {
                    total = value
                }
            val isZero: Boolean
                get() = total == 0
        }

        fun main() {
            val t = Timer()
            println(t.isZero)
            t.seconds = 3
            println(t.seconds)
            println(t.isZero)
        }
    "#);
    assert_eq!(out, &["true", "3", "false"]);
}
