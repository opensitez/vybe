//! Package + import semantics (JLS §7): package declarations,
//! single-type imports, type-import-on-demand (star), static imports,
//! fully-qualified use, and the implicit `java.lang.*` import.

use crate::helpers::run_prints;

#[test]
fn single_type_import_binds_simple_name() {
    let out = run_prints(
        r#"import java.util.HashMap;
public class Main {
    public static void main(String[] args) {
        HashMap<String, Integer> m = new HashMap<String, Integer>();
        m.put("a", 1);
        System.out.println(m.get("a"));
    }
}"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn on_demand_import_binds_simple_names() {
    let out = run_prints(
        r#"import java.util.*;
public class Main {
    public static void main(String[] args) {
        ArrayList<Integer> list = new ArrayList<Integer>();
        list.add(7);
        list.add(9);
        System.out.println(list.size());
        System.out.println(list.get(1));
    }
}"#,
    );
    assert_eq!(out, vec!["2", "9"]);
}

#[test]
fn multiple_single_type_imports() {
    let out = run_prints(
        r#"import java.util.ArrayList;
import java.util.HashMap;
public class Main {
    public static void main(String[] args) {
        ArrayList<String> list = new ArrayList<String>();
        list.add("x");
        HashMap<String, String> map = new HashMap<String, String>();
        map.put("k", "v");
        System.out.println(list.size());
        System.out.println(map.get("k"));
    }
}"#,
    );
    assert_eq!(out, vec!["1", "v"]);
}

#[test]
fn fully_qualified_type_without_import() {
    // JLS §6.7: a canonical name is usable with no import at all.
    let out = run_prints(
        r#"public class Main {
    public static void main(String[] args) {
        java.util.ArrayList list = new java.util.ArrayList();
        list.add("x");
        System.out.println(list.size());
    }
}"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn package_declaration_compiles_and_runs() {
    // JLS §7.4: package declaration heads the compilation unit.
    let out = run_prints(
        r#"package com.example.app;
public class Main {
    public static void main(String[] args) {
        System.out.println("pkg ok");
    }
}"#,
    );
    assert_eq!(out, vec!["pkg ok"]);
}

#[test]
fn package_declaration_with_imports() {
    let out = run_prints(
        r#"package com.example.app;
import java.util.ArrayList;
public class Main {
    public static void main(String[] args) {
        ArrayList<Integer> list = new ArrayList<Integer>();
        list.add(3);
        System.out.println(list.get(0));
    }
}"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn java_lang_is_implicitly_imported() {
    // JLS §7.3: every compilation unit implicitly imports java.lang.*.
    let out = run_prints(
        r#"public class Main {
    public static void main(String[] args) {
        String s = String.valueOf(12);
        System.out.println(s);
        System.out.println(Math.max(2, 9));
    }
}"#,
    );
    assert_eq!(out, vec!["12", "9"]);
}

#[test]
fn static_import_single_member() {
    // JLS §7.5.3: `import static java.lang.Math.max;` binds the bare
    // member name.
    let out = run_prints(
        r#"import static java.lang.Math.max;
public class Main {
    public static void main(String[] args) {
        System.out.println(max(2, 9));
    }
}"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn static_import_on_demand() {
    // JLS §7.5.4: `import static java.lang.Math.*;` binds every static
    // member's bare name.
    let out = run_prints(
        r#"import static java.lang.Math.*;
public class Main {
    public static void main(String[] args) {
        System.out.println(abs(-5));
    }
}"#,
    );
    assert_eq!(out, vec!["5"]);
}
