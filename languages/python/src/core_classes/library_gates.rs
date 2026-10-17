//! Immutable substring gates. The matcher retains only byte transitions and
//! ordered gate identities; it never stores program scopes or declaration ASTs.
struct Node {
    next: Vec<usize>,
    fail: usize,
    outputs: Vec<usize>,
}
pub(super) struct Matcher {
    alphabet: [usize; 256],
    nodes: Vec<Node>,
    row_count: usize,
}
impl Matcher {
    pub(super) fn new(gates: &[&str]) -> Self {
        let mut alphabet = [0; 256];
        let mut width = 1;
        for &m in gates {
            for b in m.bytes() {
                if alphabet[b as usize] == 0 {
                    alphabet[b as usize] = width;
                    width += 1;
                }
            }
        }
        let blank = || Node {
            next: vec![0; width],
            fail: 0,
            outputs: Vec::new(),
        };
        let mut nodes = vec![blank()];
        for (row, m) in gates.iter().enumerate() {
            let mut state = 0;
            for b in m.bytes() {
                let c = alphabet[b as usize];
                if nodes[state].next[c] == 0 {
                    let index = nodes.len();
                    nodes.push(blank());
                    nodes[state].next[c] = index;
                }
                state = nodes[state].next[c];
            }
            nodes[state].outputs.push(row);
        }
        let mut queue = std::collections::VecDeque::new();
        for c in 0..width {
            if nodes[0].next[c] != 0 {
                queue.push_back(nodes[0].next[c]);
            }
        }
        while let Some(state) = queue.pop_front() {
            for c in 0..width {
                let next = nodes[state].next[c];
                let fallback = nodes[nodes[state].fail].next[c];
                if next == 0 {
                    nodes[state].next[c] = fallback;
                } else {
                    nodes[next].fail = fallback;
                    let outputs = nodes[fallback].outputs.clone();
                    nodes[next].outputs.extend(outputs);
                    queue.push_back(next);
                }
            }
        }
        Self {
            alphabet,
            nodes,
            row_count: gates.len(),
        }
    }
    pub(super) fn rows(&self, source: &str) -> Vec<usize> {
        let mut matched = Vec::new();
        let mut rows = Vec::new();
        // Most source text names no library gate. Allocate bookkeeping only
        // after an actual match, and recover table order after scanning.
        let mut note = |row: usize| {
            if matched.is_empty() {
                matched.resize(self.row_count, false);
            }
            if !matched[row] {
                matched[row] = true;
                rows.push(row);
            }
        };
        for &row in &self.nodes[0].outputs {
            note(row);
        }
        let mut state = 0;
        for b in source.bytes() {
            state = self.nodes[state].next[self.alphabet[b as usize]];
            for &row in &self.nodes[state].outputs {
                note(row);
            }
        }
        rows.sort_unstable();
        rows
    }
}
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn suffixes_overlap_duplicates_empty_and_utf8_match_contains() {
        let gates = ["a", "ba", "aba", "é", "世界", "", "ba", "io", "asyncio"];
        let matcher = Matcher::new(&gates);
        let atoms = ["", "a", "b", "é", "世界", "io", "asyncio", "🙂"];
        for left in atoms {
            for middle in atoms {
                for right in atoms {
                    let source = format!("{left}{middle}{right}");
                    let expected = gates
                        .iter()
                        .enumerate()
                        .filter_map(|(index, gate)| source.contains(gate).then_some(index))
                        .collect::<Vec<_>>();
                    assert_eq!(matcher.rows(&source), expected, "{source:?}");
                }
            }
        }
    }
}
