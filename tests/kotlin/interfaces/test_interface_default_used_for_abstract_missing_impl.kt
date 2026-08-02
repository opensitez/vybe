// vybe-test: kotlin/interfaces/test_interface_default_used_for_abstract_missing_impl
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Protocol {
            fun route(): String = "default-route"
            fun name(): String
        }

        class Service : Protocol {
            override fun name(): String = "svc"
        }

        class OverrideService : Protocol {
            override fun route(): String = "custom-route"
            override fun name(): String = "custom-svc"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Protocol = Service()
            val overrideSvc: Protocol = OverrideService()
            __check((base.route()).toString(), "default-route")
            __check((base.name()).toString(), "svc")
            __check((overrideSvc.route()).toString(), "custom-route")
        }
