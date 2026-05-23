use super::helpers::run_js;

// ── Symbol.iterator ───────────────────────────────────────
#[test]
fn symbol_iterator_custom_iterable() {
    assert_eq!(run_js(r#"
const range = {
  from: 1, to: 5,
  [Symbol.iterator]() {
    let cur = this.from;
    const last = this.to;
    return {
      next() {
        return cur <= last ? { value: cur++, done: false } : { value: undefined, done: true };
      }
    };
  }
};
const result = [...range];
console.log(result.join(","));
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn symbol_iterator_used_in_for_of() {
    assert_eq!(run_js(r#"
const steps = {
  [Symbol.iterator]() {
    let n = 0;
    return { next() { return n < 3 ? { value: n++, done: false } : { done: true }; } };
  }
};
const out = [];
for (const s of steps) out.push(s);
console.log(out.join(","));
"#), vec!["0,1,2"]);
}

// ── Symbol.toPrimitive ────────────────────────────────────
#[test]
fn symbol_toprimitive_number_hint() {
    assert_eq!(run_js(r#"
const obj = {
  [Symbol.toPrimitive](hint) {
    if (hint === "number") return 42;
    if (hint === "string") return "hello";
    return true;
  }
};
console.log(+obj);
console.log(`${obj}`);
"#), vec!["42", "hello"]);
}

#[test]
fn symbol_toprimitive_default_hint() {
    assert_eq!(run_js(r#"
const token = {
  value: 100,
  [Symbol.toPrimitive](hint) {
    return hint === "string" ? "token" : this.value;
  }
};
console.log(token + 5);
"#), vec!["105"]);
}

// ── Symbol.toStringTag ────────────────────────────────────
#[test]
fn symbol_tostringtag_custom_class() {
    assert_eq!(run_js(r#"
class MyBuffer {
  get [Symbol.toStringTag]() { return "MyBuffer"; }
}
const buf = new MyBuffer();
console.log(Object.prototype.toString.call(buf));
"#), vec!["[object MyBuffer]"]);
}

#[test]
fn symbol_tostringtag_on_plain_object() {
    assert_eq!(run_js(r#"
const obj = { [Symbol.toStringTag]: "CustomTag" };
console.log(Object.prototype.toString.call(obj));
"#), vec!["[object CustomTag]"]);
}

// ── Symbol.hasInstance ────────────────────────────────────
#[test]
fn symbol_hasinstance_custom_instanceof() {
    assert_eq!(run_js(r#"
class EvenNumber {
  static [Symbol.hasInstance](val) {
    return typeof val === "number" && val % 2 === 0;
  }
}
console.log(4 instanceof EvenNumber);
console.log(3 instanceof EvenNumber);
"#), vec!["true", "false"]);
}

// ── Symbol.isConcatSpreadable ─────────────────────────────
#[test]
fn symbol_isconcatspreadable_array_like() {
    assert_eq!(run_js(r#"
const arrayLike = { 0: "a", 1: "b", length: 2, [Symbol.isConcatSpreadable]: true };
const result = ["x"].concat(arrayLike);
console.log(result.join(","));
"#), vec!["x,a,b"]);
}

#[test]
fn symbol_isconcatspreadable_false_prevents_spread() {
    assert_eq!(run_js(r#"
const arr = [1, 2];
arr[Symbol.isConcatSpreadable] = false;
const result = [0].concat(arr);
console.log(result.length);
"#), vec!["2"]);
}

// ── Symbol.species ────────────────────────────────────────
#[test]
fn symbol_species_map_returns_correct_type() {
    assert_eq!(run_js(r#"
class PowerArray extends Array {
  static get [Symbol.species]() { return Array; }
}
const arr = new PowerArray(1, 2, 3);
const mapped = arr.map(x => x * 2);
console.log(mapped instanceof PowerArray);
console.log(mapped instanceof Array);
"#), vec!["false", "true"]);
}

// ── Symbol.match ──────────────────────────────────────────
#[test]
fn symbol_match_custom_matcher() {
    assert_eq!(run_js(r#"
const matcher = {
  [Symbol.match](str) {
    return str.startsWith("hello") ? ["hello"] : null;
  }
};
const result = "hello world".match(matcher);
console.log(result[0]);
"#), vec!["hello"]);
}

// ── Symbol.replace ────────────────────────────────────────
#[test]
fn symbol_replace_custom_replacer() {
    assert_eq!(run_js(r#"
const replacer = {
  [Symbol.replace](str, replacement) {
    return str.split("x").join(replacement);
  }
};
console.log("axbxc".replace(replacer, "-"));
"#), vec!["a-b-c"]);
}

// ── Symbol.search ─────────────────────────────────────────
#[test]
fn symbol_search_custom_search() {
    assert_eq!(run_js(r#"
const searcher = {
  [Symbol.search](str) { return str.indexOf("x"); }
};
console.log("abxcd".search(searcher));
"#), vec!["2"]);
}

// ── Symbol.split ─────────────────────────────────────────
#[test]
fn symbol_split_custom_splitter() {
    assert_eq!(run_js(r#"
const splitter = {
  [Symbol.split](str) {
    return str.split("").filter(c => c !== " ");
  }
};
const result = "a b c".split(splitter);
console.log(result.join(","));
"#), vec!["a,b,c"]);
}

// ── Symbol basics ─────────────────────────────────────────
#[test]
fn symbol_unique_per_call() {
    assert_eq!(run_js(r#"
const s1 = Symbol("id");
const s2 = Symbol("id");
console.log(s1 === s2);
console.log(typeof s1);
"#), vec!["false", "symbol"]);
}

#[test]
fn symbol_description_property() {
    assert_eq!(run_js(r#"
const s = Symbol("myDescription");
console.log(s.description);
"#), vec!["myDescription"]);
}

#[test]
fn symbol_tostring() {
    assert_eq!(run_js(r#"
const s = Symbol("test");
console.log(s.toString());
"#), vec!["Symbol(test)"]);
}

#[test]
fn symbol_for_global_registry() {
    assert_eq!(run_js(r#"
const s1 = Symbol.for("shared");
const s2 = Symbol.for("shared");
console.log(s1 === s2);
"#), vec!["true"]);
}

#[test]
fn symbol_keyfor_returns_key() {
    assert_eq!(run_js(r#"
const s = Symbol.for("myKey");
console.log(Symbol.keyFor(s));
"#), vec!["myKey"]);
}

#[test]
fn symbol_keyfor_local_symbol_is_undefined() {
    assert_eq!(run_js(r#"
const s = Symbol("local");
console.log(Symbol.keyFor(s) === undefined);
"#), vec!["true"]);
}

#[test]
fn symbol_as_object_key() {
    assert_eq!(run_js(r#"
const id = Symbol("id");
const user = { [id]: 42, name: "Alice" };
console.log(user[id]);
console.log(user.name);
"#), vec!["42", "Alice"]);
}

#[test]
fn symbol_not_in_json_stringify() {
    assert_eq!(run_js(r#"
const sym = Symbol("hidden");
const obj = { [sym]: "secret", visible: "yes" };
const json = JSON.stringify(obj);
console.log(json);
"#), vec![r#"{"visible":"yes"}"#]);
}

#[test]
fn symbol_not_in_object_keys() {
    assert_eq!(run_js(r#"
const sym = Symbol("x");
const obj = { [sym]: 1, a: 2 };
console.log(Object.keys(obj).join(","));
"#), vec!["a"]);
}

#[test]
fn symbol_in_object_getownpropertysymbols() {
    assert_eq!(run_js(r#"
const sym = Symbol("x");
const obj = { [sym]: 1, a: 2 };
const syms = Object.getOwnPropertySymbols(obj);
console.log(syms.length);
"#), vec!["1"]);
}

#[test]
fn symbol_iterator_string_default() {
    assert_eq!(run_js(r#"
const chars = [...'abc'];
console.log(chars.join("-"));
"#), vec!["a-b-c"]);
}

#[test]
fn symbol_iterator_map_default() {
    assert_eq!(run_js(r#"
const m = new Map([["a", 1], ["b", 2]]);
const pairs = [...m];
console.log(pairs.length);
"#), vec!["2"]);
}

#[test]
fn symbol_iterator_set_default() {
    assert_eq!(run_js(r#"
const s = new Set([10, 20, 30]);
const vals = [...s];
console.log(vals.join(","));
"#), vec!["10,20,30"]);
}

#[test]
fn symbol_asynciterator_protocol() {
    assert_eq!(run_js(r#"
async function collect() {
  const results = [];
  const asyncIterable = {
    [Symbol.asyncIterator]() {
      let i = 0;
      return {
        async next() {
          if (i < 3) return { value: i++, done: false };
          return { value: undefined, done: true };
        }
      };
    }
  };
  for await (const val of asyncIterable) results.push(val);
  console.log(results.join(","));
}
collect();
"#), vec!["0,1,2"]);
}

#[test]
fn symbol_toprimitive_all_hints() {
    assert_eq!(run_js(r#"
const obj = {
  [Symbol.toPrimitive](hint) { return hint; }
};
console.log(String(obj));
console.log(Number(obj) === 0);
"#), vec!["string", "false"]);
}

#[test]
fn symbol_well_known_iterator_protocol_array() {
    assert_eq!(run_js(r#"
const iter = [1, 2, 3][Symbol.iterator]();
console.log(iter.next().value);
console.log(iter.next().value);
console.log(iter.next().done);
"#), vec!["1", "2", "false"]);
}
