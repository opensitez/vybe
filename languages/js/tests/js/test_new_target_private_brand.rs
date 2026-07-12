crate::js_cases! {
    new_target_is_undefined_in_plain_call => {
        r#"
function kind() {
  return new.target === undefined ? "plain" : "construct";
}
console.log(kind());
"#,
        ["plain"]
    };

    new_target_points_at_called_constructor => {
        r#"
function Person() {
  console.log(new.target === Person);
}
new Person();
"#,
        ["true"]
    };

    new_target_flows_through_super_constructor => {
        r#"
class Base {
  constructor() {
    console.log(new.target === Derived);
  }
}
class Derived extends Base {}
new Derived();
"#,
        ["true"]
    };

    new_target_is_captured_by_arrow_in_constructor => {
        r#"
function Box() {
  const read = () => new.target;
  console.log(read() === Box);
}
new Box();
"#,
        ["true"]
    };

    private_brand_check_accepts_own_instance => {
        r#"
class Counter {
  #count = 0;
  hasBrand(value) {
    return #count in value;
  }
}
const counter = new Counter();
console.log(counter.hasBrand(counter));
"#,
        ["true"]
    };

    private_brand_check_rejects_plain_object => {
        r#"
class Counter {
  #count = 0;
  hasBrand(value) {
    return #count in value;
  }
}
console.log(new Counter().hasBrand({ count: 0 }));
"#,
        ["false"]
    };

    private_brand_check_accepts_subclass_instance => {
        r#"
class Base {
  #token = 1;
  hasBrand(value) {
    return #token in value;
  }
}
class Derived extends Base {}
console.log(new Base().hasBrand(new Derived()));
"#,
        ["true"]
    };

    private_brand_checks_use_distinct_brands_per_class => {
        r#"
class Left {
  #value = 1;
  hasBrand(value) {
    return #value in value;
  }
}
class Right {
  #value = 1;
}
console.log(new Left().hasBrand(new Right()));
"#,
        ["false"]
    };
}
