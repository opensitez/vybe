// vybe-test: kotlin/visibility/test_private_members_are_not_accessible_outside_declaring_class
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Item {
            private val secret: Int = 9
        }

        fun main() {
            val item = Item()
            println(item.secret)
        }

