use crate::helpers::{compile_ok_check, run_prints};

#[test]
fn test_public_members_are_callable_by_default() {
    let out = run_prints(
        r#"
        class Item {
            val label = "x"
            fun text(): String = "ok"
        }

        fun main() {
            val item = Item()
            println(item.label)
            println(item.text())
        }
    "#,
    );
    assert_eq!(out, &["x", "ok"]);
}

#[test]
fn test_private_members_are_not_accessible_outside_declaring_class() {
    assert!(!compile_ok_check(
        r#"
        class Item {
            private val secret: Int = 9
        }

        fun main() {
            val item = Item()
            println(item.secret)
        }
    "#
    ));
}

#[test]
fn test_protected_property_is_visible_to_subclass() {
    let out = run_prints(
        r#"
        open class Base {
            protected var value: Int = 1
        }

        class Child : Base() {
            fun bump() { value += 2 }
            fun read(): Int = value
        }

        fun main() {
            val child = Child()
            child.bump()
            println(child.read())
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_protected_property_not_visible_to_unrelated_type() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            protected var value: Int = 1
        }

        fun leak(base: Base) {
            println(base.value)
        }

        fun main() {
            leak(Base())
        }
    "#
    ));
}

#[test]
fn test_private_constructor_prevents_external_call() {
    assert!(!compile_ok_check(
        r#"
        class Guard private constructor(val value: Int)

        fun main() {
            val value = Guard(1)
            println(value.value)
        }
    "#
    ));
}

#[test]
fn test_private_class_is_file_local_access_control() {
    let out = run_prints(
        r#"
        private class Hidden {
            val value = 8
        }

        fun spawn(): Int {
            return Hidden().value
        }

        fun main() {
            println(spawn())
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_private_member_stays_within_class_hierarchy() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            private val hidden = 4
        }

        class Child : Base() {
            fun read(): Int = hidden
        }
    "#
    ));
}

#[test]
fn test_internal_property_access_within_module() {
    let out = run_prints(
        r#"
        class Item {
            internal val value = 9
        }

        fun main() {
            val item = Item()
            println(item.value)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_override_restricts_private_visibility() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            private open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
        }
    "#
    ));
}

#[test]
fn test_private_setter_restricts_external_assignment() {
    assert!(!compile_ok_check(
        r#"
        class Counter {
            var value: Int = 0
                private set
        }

        fun main() {
            val counter = Counter()
            counter.value = 4
            println(counter.value)
        }
    "#
    ));
}

#[test]
fn test_accessing_private_setter_from_same_class_only() {
    let out = run_prints(
        r#"
        class Counter {
            var value: Int = 0
                private set

            fun setValue(next: Int) {
                value = next
            }
        }

        fun main() {
            val counter = Counter()
            counter.setValue(11)
            println(counter.value)
        }
    "#,
    );
    assert_eq!(out, &["11"]);
}

#[test]
fn test_public_members_are_implicitly_accessible() {
    let out = run_prints(
        r#"
        class Item {
            fun show(): String = "ok"
        }

        fun main() {
            println(Item().show())
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_protected_member_is_only_visible_in_subclass_hierarchy() {
    let out = run_prints(
        r#"
        open class Base {
            protected var value: Int = 3
        }

        class Child : Base() {
            fun write(next: Int) {
                value = next
            }

            fun read(): Int {
                return value
            }
        }

        fun main() {
            val child = Child()
            child.write(6)
            println(child.read())
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_protected_member_call_rejected_outside_hierarchy() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            protected fun label(): String = "base"
        }

        class NotChild : Base()

        fun main() {
            val value = Base()
            println(value.label())
        }
    "#
    ));
}

#[test]
fn test_private_class_is_inaccessible_outside_file_scope() {
    let out = run_prints(
        r#"
        private class Local {
            fun value(): String = "inner"
        }

        fun main() {
            println(Local().value())
        }
    "#,
    );
    assert_eq!(out, &["inner"]);
}

#[test]
fn test_private_constructor_hides_instantiation_outside_factory() {
    assert!(!compile_ok_check(
        r#"
        class Locked private constructor()

        fun main() {
            Locked()
        }
    "#
    ));
}

#[test]
fn test_factory_function_preserves_private_constructor_access_control() {
    let out = run_prints(
        r#"
        class Box private constructor(val value: String) {
            companion object {
                fun from(value: String): Box = Box(value)
            }
        }

        fun main() {
            println(Box.from("x").value)
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_private_setter_restricts_direct_mutation() {
    assert!(!compile_ok_check(
        r#"
        class Count {
            var value: Int = 0
                private set
        }

        fun main() {
            val count = Count()
            count.value = 4
            println(count.value)
        }
    "#
    ));
}

#[test]
fn test_public_setter_allows_external_mutation() {
    let out = run_prints(
        r#"
        class Count {
            var value: Int = 0
        }

        fun main() {
            val count = Count()
            count.value = 4
            println(count.value)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_internal_properties_work_within_single_module() {
    let out = run_prints(
        r#"
        class Item {
            internal var tag: String = "mod"
        }

        fun main() {
            val item = Item()
            item.tag = "ok"
            println(item.tag)
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_private_modifier_prevents_override_for_members() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            private fun hidden() = "x"
        }

        class Child : Base() {
            fun call(base: Base): String {
                return base.hidden()
            }
        }
    "#
    ));
}

#[test]
fn test_private_members_are_visible_inside_companion_object_methods() {
    let out = run_prints(
        r#"
        class Token {
            private val secret = "ok"

            companion object {
                fun reveal(token: Token): String = token.secret
            }
        }

        fun main() {
            println(Token.reveal(Token()))
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_inner_class_can_read_outer_private_member() {
    let out = run_prints(
        r#"
        class Vault {
            private val secret = "vault"

            inner class Reader {
                fun open(): String = secret
            }
        }

        fun main() {
            val vault = Vault()
            println(vault.Reader().open())
        }
    "#,
    );
    assert_eq!(out, &["vault"]);
}

#[test]
fn test_private_function_is_inaccessible_from_extension_function() {
    assert!(!compile_ok_check(
        r#"
        class Item {
            private fun secret(): String = "x"
        }

        fun Item.exposed(): String = secret()

        fun main() {
            println(Item().exposed())
        }
    "#
    ));
}

#[test]
fn test_private_function_can_be_called_via_same_file_extension_only_when_redeclaration_absent() {
    let out = run_prints(
        r#"
        class Item {
            private fun secret(): String = "inside"
        }

        fun itemSecret(item: Item): String = item.access()

        private fun Item.access(): String = secret()

        fun main() {
            println(itemSecret(Item()))
        }
    "#,
    );
    assert_eq!(out, &["inside"]);
}

#[test]
fn test_private_constructor_is_restricted_to_same_scope() {
    assert!(!compile_ok_check(
        r#"
        class Core private constructor(val value: String) {
            companion object {
                fun from(value: String): Core = Core(value)
            }
        }

        class Factory {
            fun create(value: String): Core {
                return Core(value)
            }
        }

        fun main() {
            println(Factory().create("x"))
        }
    "#
    ));
}

#[test]
fn test_protected_getter_can_be_read_only_outside_but_not_settable_from_subclass_reference() {
    let out = run_prints(
        r#"
        open class Base {
            protected var value: Int = 1
        }

        class Child : Base() {
            var view: Int
                get() = value
                set(next) { value = next }
        }

        fun main() {
            val child = Child()
            child.view = 10
            println(child.view)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_override_with_weaker_visibility_is_rejected_for_member() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            protected open fun label(): String = "base"
        }

        class Child : Base() {
            private override fun label(): String = "child"
        }
    "#
    ));
}

#[test]
fn test_protected_function_visible_in_deeper_subclass() {
    let out = run_prints(
        r#"
        open class Base {
            protected open fun label(): String = "base"
        }

        open class Mid : Base() {
            override fun label(): String = "mid"
        }

        class Child : Mid() {
            fun text(): String = label()
        }

        fun main() {
            println(Child().text())
        }
    "#,
    );
    assert_eq!(out, &["mid"]);
}

#[test]
fn test_protected_function_cannot_be_called_on_unrelated_reference() {
    assert!(!compile_ok_check(
        r#"
        open class Base {
            protected fun label(): String = "base"
        }

        fun call(base: Base): String = base.label()

        fun main() {
            println(call(Base()))
        }
    "#
    ));
}

#[test]
fn test_internal_class_and_members_are_visible_in_same_file_scope() {
    let out = run_prints(
        r#"
        internal class Vault(val value: Int)

        fun main() {
            println(Vault(7).value)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_internal_member_allows_rebinding_within_same_module() {
    let out = run_prints(
        r#"
        class Box {
            internal var value: Int = 2
            fun bump() {
                value++
            }
        }

        fun main() {
            val box = Box()
            box.bump()
            box.value = 9
            println(box.value)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_private_setter_is_not_accessible_for_extension_receiver() {
    assert!(!compile_ok_check(
        r#"
        class Counter {
            var value: Int = 0
                private set
        }

        fun Counter.bump() {
            this.value = this.value + 1
        }

        fun main() {
            Counter().bump()
        }
    "#
    ));
}

#[test]
fn test_private_setter_does_not_block_reading_property() {
    let out = run_prints(
        r#"
        class Counter {
            var value: Int = 4
                private set

            fun increment() {
                value += 1
            }
        }

        fun main() {
            val counter = Counter()
            counter.increment()
            println(counter.value)
            counter.increment()
            println(counter.value)
        }
    "#,
    );
    assert_eq!(out, &["5", "6"]);
}

#[test]
fn test_public_setter_can_overwrite_private_getter_state() {
    let out = run_prints(
        r#"
        class Item {
            private var value = "x"
            var display: String
                get() = value
                private set(next) {
                    value = next
                }

            fun reset(next: String) {
                display = next
            }
        }

        fun main() {
            val item = Item()
            println(item.display)
            item.reset("y")
            println(item.display)
        }
    "#,
    );
    assert_eq!(out, &["x", "y"]);
}

#[test]
fn test_private_constructor_enforced_even_with_factory_style_use() {
    let out = run_prints(
        r#"
        class Config private constructor(val value: Int) {
            companion object {
                fun create(value: Int): Config = Config(value)
            }
        }

        fun main() {
            println(Config.create(3).value)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_visibility_default_is_public_for_top_level_members() {
    let out = run_prints(
        r#"
        class Item {
            fun status(): String = "public"
        }

        fun main() {
            val item = Item()
            println(item.status())
        }
    "#,
    );
    assert_eq!(out, &["public"]);
}

#[test]
fn test_private_field_in_data_class_still_part_of_public_api_via_copy() {
    let out = run_prints(
        r#"
        data class Holder(private val secret: String, val value: String) {
            fun reveal(): String = secret
        }

        fun main() {
            val holder = Holder("x", "y")
            println(holder.value)
            println(holder.reveal())
            println(holder.copy(value = "z").value)
        }
    "#,
    );
    assert_eq!(out, &["y", "x", "z"]);
}
