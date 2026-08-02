// vybe-test: kotlin/inheritance_dispatch/test_interface_default_with_class_super_resolution
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Tracer {
            fun route(): String = "trace"
        }

        open class Base {
            open fun route(): String = "base"
        }

        class Logger : Base(), Tracer {
            override fun route(): String = super<Tracer>.route() + ":" + super<Base>.route()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val logger: Base = Logger()
            __check((logger.route()).toString(), "trace:base")
            __check(((logger as Logger).route()).toString(), "trace:base")
        }
