// vybe-test: kotlin/smart_casts/test_nested_smart_cast_after_outer_check
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

open class Base
        class Holder(val text: String) : Base()
        class Wrapper(val child: Base)

        fun main() {
            val value: Any = Wrapper(Holder("ok"))
            if (value is Wrapper && value.child is Holder) {
                println(value.child.text)
            } else {
                println("no")
            }
        }

