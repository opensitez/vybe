/// Generator state machines — traffic lights, parsers, coroutines, backtracking
use super::helpers::run_js;

#[test]
fn traffic_light_state_machine() {
    assert_eq!(
        run_js(
            r#"
function* trafficLight() {
    while (true) {
        yield "red";
        yield "green";
        yield "yellow";
    }
}
const light = trafficLight();
const states = [];
for (let i = 0; i < 7; i++) states.push(light.next().value);
console.log(states.join(","));
"#
        ),
        vec!["red,green,yellow,red,green,yellow,red"]
    );
}

#[test]
fn csv_parser_generator() {
    assert_eq!(
        run_js(
            r#"
function* parseCSV(text) {
    const lines = text.split("\n").filter(Boolean);
    for (const line of lines) {
        yield line.split(",").map(s => s.trim());
    }
}
const csv = "Alice,30,Engineer\nBob,25,Designer\nCharlie,35,Manager";
const rows = [...parseCSV(csv)];
console.log(rows.length);
console.log(rows[0][0]);
console.log(rows[1][2]);
"#
        ),
        vec!["3", "Alice", "Designer"]
    );
}

#[test]
fn token_lexer_generator() {
    assert_eq!(
        run_js(
            r#"
function* tokenize(expr) {
    const re = /\d+|[+\-*/()]/g;
    let m;
    while ((m = re.exec(expr)) !== null) {
        yield m[0];
    }
}
const tokens = [...tokenize("1 + 2 * (3 - 4)")];
console.log(tokens.join(","));
"#
        ),
        vec!["1,+,2,*,(,3,-,4,)"]
    );
}

#[test]
fn generator_as_coroutine_send() {
    assert_eq!(
        run_js(
            r#"
function* counter(start = 0) {
    let n = start;
    while (true) {
        const reset = yield n;
        if (reset === true) n = start;
        else n++;
    }
}
const gen = counter(10);
gen.next(); // start
console.log(gen.next().value);  // 11
console.log(gen.next().value);  // 12
console.log(gen.next(true).value); // reset to 10
console.log(gen.next().value);  // 11
"#
        ),
        vec!["11", "12", "10", "11"]
    );
}

#[test]
fn round_robin_scheduler() {
    assert_eq!(
        run_js(
            r#"
function* roundRobin(tasks) {
    const generators = tasks.map(t => t());
    while (generators.length > 0) {
        for (let i = generators.length - 1; i >= 0; i--) {
            const result = generators[i].next();
            if (result.done) generators.splice(i, 1);
            else yield result.value;
        }
    }
}
function* task(name, steps) {
    for (let i = 0; i < steps; i++) yield `${name}:${i}`;
}
const log = [...roundRobin([
    () => task("A", 2),
    () => task("B", 2),
])];
console.log(log.join(","));
"#
        ),
        vec!["B:0,A:0,B:1,A:1"]
    );
}

#[test]
fn depth_first_search_generator() {
    assert_eq!(
        run_js(
            r#"
function* dfs(graph, start, visited = new Set()) {
    if (visited.has(start)) return;
    visited.add(start);
    yield start;
    for (const neighbor of (graph[start] || [])) {
        yield* dfs(graph, neighbor, visited);
    }
}
const graph = {
    A: ["B", "C"],
    B: ["D"],
    C: ["D", "E"],
    D: [],
    E: []
};
const order = [...dfs(graph, "A")];
console.log(order.join(","));
"#
        ),
        vec!["A,B,D,C,E"]
    );
}

#[test]
fn generator_as_observable() {
    assert_eq!(
        run_js(
            r#"
function* events(data) {
    for (const item of data) {
        if (item > 0) yield { type: "positive", value: item };
        else if (item < 0) yield { type: "negative", value: item };
        else yield { type: "zero", value: 0 };
    }
}
const evts = [...events([1, -2, 0, 3])];
console.log(evts.map(e => e.type).join(","));
console.log(evts.map(e => e.value).join(","));
"#
        ),
        vec!["positive,negative,zero,positive", "1,-2,0,3"]
    );
}

#[test]
fn generator_state_machine_early_break_cleanup() {
    assert_eq!(
        run_js(
            r#"
function* stateMachine() {
    try {
        yield "state1";
        yield "state2";
    } finally {
        console.log("cleaned_up");
    }
}
for (const s of stateMachine()) {
    console.log(s);
    break;
}
"#
        ),
        vec!["state1", "cleaned_up"]
    );
}

