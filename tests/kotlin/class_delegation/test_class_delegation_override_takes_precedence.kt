// vybe-test: kotlin/class_delegation/test_class_delegation_override_takes_precedence
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Service { fun name(): String }

        class Primary : Service {
            override fun name() = "primary"
        }

        class Decorated(delegate: Service) : Service by delegate {
            override fun name() = "decorated"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Decorated(Primary()).name()).toString(), "decorated")
        }
