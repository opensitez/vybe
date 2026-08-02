// vybe-test: kotlin/interfaces/test_interface_declaration
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Printable {
            fun printMsg()
        }

        class MessagePrinter : Printable {
            fun printMsg() {
                __check(("interface message").toString(), "interface message")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val printer = MessagePrinter()
            printer.printMsg()
        }
