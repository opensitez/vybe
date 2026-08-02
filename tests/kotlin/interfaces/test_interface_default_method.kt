// vybe-test: kotlin/interfaces/test_interface_default_method
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Notifier {
            fun notify(): String {
                return "notified"
            }
        }

        class SilentNotifier : Notifier

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n: Notifier = SilentNotifier()
            __check((n.notify()).toString(), "notified")
        }
