// vybe-test: kotlin/class_delegation/test_simple_class_delegation_forwards_call
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Greeter { fun hello(): String }

        class Base(private val name: String) : Greeter {
            override fun hello() = "hello:$name"
        }

        class Wrapper(delegate: Greeter) : Greeter by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w = Wrapper(Base("kotlin"))
            __check((w.hello()).toString(), "hello:kotlin")
        }
