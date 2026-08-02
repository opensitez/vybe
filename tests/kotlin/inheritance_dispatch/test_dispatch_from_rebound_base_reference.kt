// vybe-test: kotlin/inheritance_dispatch/test_dispatch_from_rebound_base_reference
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Printer {
            open fun emit(prefix: String): String = prefix + ":base"
        }

        class Loud : Printer() {
            override fun emit(prefix: String): String = prefix + ":loud"
        }

        fun emitTwice(printer: Printer): String {
            return printer.emit("x") + "," + printer.emit("y")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var printer: Printer = Printer()
            printer = Loud()
            __check((emitTwice(printer)).toString(), "x:loud,y:loud")
        }
