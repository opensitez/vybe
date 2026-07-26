use super::helpers::run_prints;

fn assert_int(source: &str, expected: i64) {
    assert_eq!(run_prints(source), vec![expected.to_string()]);
}

#[test]
fn php_oop_class_shapes() {
    for idx in 1..=140_i64 {
        let class_name = format!("Model{idx}");
        let source = format!(
            "<?php\nclass {class_name} {{\n    public function __construct(public int $id, public int $seed) {{}}\n\n    public function value(): int {{\n        return $this->id + $this->seed;\n    }}\n}}\n\n$instance = new {class_name}({idx}, {idx});\necho $instance->value();\n",
        );
        assert_int(&source, idx + idx);

        let base_name = format!("Base{idx}");
        let child_name = format!("Child{idx}");
        let inheritance_source = format!(
            "<?php\nclass {base_name} {{\n    public function score(int $offset): int {{\n        return $offset + {idx};\n    }}\n}}\n\nclass {child_name} extends {base_name} {{\n    public function score(int $offset): int {{\n        return parent::score($offset) + {idx};\n    }}\n}}\n\necho (new {child_name}())->score(1);\n",
        );
        assert_int(&inheritance_source, idx * 2 + 1);

        let trait_name = format!("Trait{idx}");
        let trait_user = format!("TraitUser{idx}");
        let trait_source = format!(
            "<?php\ntrait {trait_name} {{\n    public function marker(): int {{\n        return {idx};\n    }}\n}}\n\nclass {trait_user} {{\n    use {trait_name};\n\n    public function value(): int {{\n        return $this->marker() + {idx};\n    }}\n}}\n\necho (new {trait_user}())->value();\n",
        );
        assert_int(&trait_source, idx * 2);

        let interface_name = format!("Iface{idx}");
        let impl_name = format!("Impl{idx}");
        let interface_source = format!(
            "<?php\ninterface {interface_name} {{\n    public function compute(int $left, int $right): int;\n}}\n\nclass {impl_name} implements {interface_name} {{\n    public function compute(int $left, int $right): int {{\n        return ($left + $right) * {idx};\n    }}\n}}\n\n$svc = new {impl_name}();\necho $svc->compute(2, 3);\n",
        );
        assert_int(&interface_source, 5 * idx);
    }
}
