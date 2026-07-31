use crate::helpers::run_prints;

#[test]
fn test_interface_default_method_is_used_when_not_overridden() {
    let out = run_prints(r#"
        interface Logger {
            fun prefix(): String = "log"
            fun format(message: String): String = prefix() + ":" + message
        }

        class DefaultLogger : Logger

        fun main() {
            println(DefaultLogger().format("ok"))
        }
    "#);
    assert_eq!(out, &["log:ok"]);
}

#[test]
fn test_interface_default_method_can_be_overridden() {
    let out = run_prints(r#"
        interface Messenger {
            fun prefix(): String = "base"
            fun format(value: String): String = prefix() + ":" + value
        }

        class DefaultMessenger : Messenger

        class LoudMessenger : Messenger {
            override fun format(value: String): String = prefix() + ":" + value.toUpperCase()
        }

        fun main() {
            println(DefaultMessenger().format("ok"))
            println(LoudMessenger().format("ok"))
        }
    "#);
    assert_eq!(out, &["base:ok", "base:OK"]);
}
