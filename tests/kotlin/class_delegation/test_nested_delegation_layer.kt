// vybe-test: kotlin/class_delegation/test_nested_delegation_layer
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Printer { fun print(value: Int): String }

        class BasePrinter : Printer {
            override fun print(value: Int): String = "base=$value"
        }

        class PrefixPrinter(delegate: Printer) : Printer by delegate
        class WrapperPrinter(delegate: Printer) : Printer by PrefixPrinter(delegate)

        fun main() {
            println(WrapperPrinter(BasePrinter()).print(4))
        }

