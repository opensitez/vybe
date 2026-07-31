use crate::helpers::run_prints;

#[test]
fn test_lazy_initialized_property_runs_once() {
    let out = run_prints(r#"
        class Holder {
            var calls = 0
            val value: Int by lazy {
                calls += 1
                10
            }
        }

        fun main() {
            val h = Holder()
            println(h.calls)
            println(h.value)
            println(h.value)
            println(h.calls)
        }
    "#);
    assert_eq!(out, &["0", "10", "10", "1"]);
}

#[test]
fn test_lazy_value_can_depend_on_previous_state() {
    let out = run_prints(r#"
        class Holder {
            var seed = 5
            val value: Int by lazy {
                seed * 2
            }
        }

        fun main() {
            val h = Holder()
            println(h.value)
            h.seed = 9
            println(h.value)
        }
    "#);
    assert_eq!(out, &["10", "10"]);
}
