fn semantic_code(
    c: &Chunk,
    imports: &[vybe_runtime::chunk::Import],
) -> Vec<(
    Op,
    Vec<u8>,
    Option<vybe_runtime::chunk::Import>,
    Option<String>,
)> {
    let mut out = vec![];
    let mut pc = 0;
    while pc < c.code.len() {
        let op = Op::decode(c.read_u16(pc), c.read_u16(pc + 2)).unwrap();
        let end = pc + 4 + op.operand_format().size_in(&c.code, pc + 4);
        let mut operands = c.code[pc + 4..end].to_vec();
        let import = if op == Op::CALL {
            let import = imports[c.read_u16(pc + 4) as usize].clone();
            operands[..2].fill(0);
            Some(import)
        } else {
            None
        };
        let global = if op == Op::GLOBAL_GET || op == Op::GLOBAL_SET {
            let index = c.read_u32(pc + 4) as usize;
            let name = if !c.globals.is_empty() {
                c.globals[index].clone()
            } else {
                let Value::String(name) = &c.constants[index] else {
                    panic!("named frozen global")
                };
                name.to_string()
            };
            operands.fill(0);
            Some(name)
        } else {
            None
        };
        out.push((op, operands, import, global));
        pc = end;
    }
    out
}
