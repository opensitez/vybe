use crate::helpers::run_prints;

#[test]
fn test_object_singleton_maintains_shared_state() {
    let out = run_prints(
        r#"
        object Config {
            var enabled = false
            fun enable() { enabled = true }
            fun isEnabled(): Boolean = enabled
        }

        fun main() {
            Config.enable()
            println(Config.isEnabled())
            Config.enabled = false
            println(Config.isEnabled())
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_object_singleton_can_be_used_as_utility() {
    let out = run_prints(
        r#"
        object Formatter {
            fun wrap(value: String): String = "<" + value + ">"
        }

        fun main() {
            println(Formatter.wrap("a"))
            println(Formatter.wrap("b"))
        }
    "#,
    );
    assert_eq!(out, &["<a>", "<b>"]);
}
