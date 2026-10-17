use super::{HashSet, SharedMemory, VM, VMError, Value};

pub(super) struct Snapshot {
    memory: SharedMemory,
    extra_memories: Vec<SharedMemory>,
    extra_memory_max_pages: Vec<Option<usize>>,
    memory_is_64: Vec<bool>,
    memory0_is_64: bool,
    active_memory: usize,
    tables: Vec<Vec<Value>>,
    table_maxes: Vec<Option<usize>>,
    table_is_64: Vec<bool>,
    table0_is_64: bool,
    data: Vec<Vec<u8>>,
    elements: Vec<Vec<Value>>,
    dropped_data: HashSet<u32>,
    dropped_elements: HashSet<u32>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Op;
    use crate::chunk::{ActiveDataSegment, ActiveElementSegment, Chunk};

    fn native_module(offset: u64) -> Vec<Chunk> {
        let mut script = Chunk::new("<script>");
        script.native_global_import_count = Some(0);
        script.memory_min_pages = vec![1];
        script.memory_max_pages = vec![Some(2)];
        script.table_min_sizes = vec![1];
        script.table_max_sizes = vec![Some(2)];
        script.data_segments = vec![vec![42, 43, 44]];
        script.elem_segments = vec![vec![Value::I32(0)]];
        script.active_data_segments = vec![ActiveDataSegment {
            memory_index: 0,
            data_index: 0,
            offset,
        }];
        script.active_elem_segments = vec![ActiveElementSegment {
            table_index: 0,
            elem_index: 0,
            offset: 0,
        }];
        script.emit_op(Op::RETURN, 0);
        let mut function = Chunk::new("native_function");
        function.emit_op(Op::RETURN, 0);
        vec![script, function]
    }

    #[test]
    fn linked_native_module_initializes_memory_tables_and_segments() {
        for linked in [false, true] {
            let mut vm = VM::new();
            let chunks = native_module(0);
            if linked {
                vm.run_linked(chunks, vec![]).unwrap();
            } else {
                vm.run(chunks).unwrap();
            }
            assert_eq!(vm.memory.len(), 65536);
            vm.memory
                .with_buffer_mut(|bytes| assert_eq!(&bytes[..3], &[42, 43, 44]));
            assert_eq!(vm.wasm_tables[0].len(), 1);
            assert!(matches!(vm.wasm_tables[0][0], Value::Object(_)));
        }
    }

    fn caller() -> VM {
        let mut vm = VM::new();
        vm.memory.resize(8, 9);
        vm.memory_is_64 = vec![true];
        vm.memory0_is_64 = true;
        vm.wasm_tables = vec![vec![Value::I32(71)]];
        vm.wasm_table_maxes = vec![Some(11)];
        vm.table_is_64 = vec![true];
        vm.table0_is_64 = true;
        vm.data_segments = vec![vec![7, 8]];
        vm.elem_segments = vec![vec![Value::I32(72)]];
        vm.dropped_data.insert(3);
        vm.dropped_elems.insert(4);
        vm
    }

    fn assert_caller(vm: &VM) {
        assert_eq!(vm.memory.len(), 8);
        vm.memory
            .with_buffer_mut(|bytes| assert_eq!(bytes, &[9; 8]));
        assert_eq!(vm.memory_is_64, vec![true]);
        assert!(vm.memory0_is_64);
        assert_eq!(vm.wasm_tables, vec![vec![Value::I32(71)]]);
        assert_eq!(vm.wasm_table_maxes, vec![Some(11)]);
        assert_eq!(vm.table_is_64, vec![true]);
        assert!(vm.table0_is_64);
        assert_eq!(vm.data_segments, vec![vec![7, 8]]);
        assert_eq!(vm.elem_segments, vec![vec![Value::I32(72)]]);
        assert!(vm.dropped_data.contains(&3));
        assert!(vm.dropped_elems.contains(&4));
    }

    #[test]
    fn nested_native_initialization_restores_caller_on_success_and_trap() {
        for offset in [0, 65536] {
            let mut vm = caller();
            let result = vm.run_linked_nested(native_module(offset), vec![]);
            assert_eq!(result.is_ok(), offset == 0);
            assert_caller(&vm);
        }
    }

    #[test]
    fn nested_native_execution_trap_restores_caller_storage() {
        let mut vm = caller();
        let mut chunks = native_module(0);
        chunks[0].code.clear();
        chunks[0].emit_op(Op::UNREACHABLE, 0);
        assert!(vm.run_linked_nested(chunks, vec![]).is_err());
        assert_caller(&vm);
    }

    #[test]
    fn source_include_does_not_reinitialize_native_storage() {
        let mut vm = caller();
        let mut source = Chunk::new("source_include");
        source.emit_op(Op::RETURN, 0);
        vm.run_linked_nested(vec![source], vec![]).unwrap();
        assert_caller(&vm);
    }

    #[test]
    fn native_segment_modes_copy_drop_and_preserve_passive_payloads() {
        use crate::chunk::{
            ConstExpr, NativeElementSegment, NativeSegmentMode as Mode, NativeSegments,
        };
        use crate::opcode::heaptype::HeapType;
        for linked in [false, true] {
            let mut chunks = native_module(0);
            chunks[0].active_data_segments.clear();
            chunks[0].active_elem_segments.clear();
            chunks[0].elem_segments.clear();
            chunks[0].data_segments.push(vec![81]);
            let element = |mode| NativeElementSegment {
                mode,
                reference_type: (HeapType::Abstract(0x70), true, false, false),
                items: vec![ConstExpr::RefFunc(1)].into(),
            };
            chunks[0].native_segments = Some(std::sync::Arc::new(NativeSegments {
                function_import_types: Vec::new(),
                global_types: Vec::new(),
                data_modes: vec![
                    Mode::Active {
                        index: 0,
                        offset: ConstExpr::Value(Value::I32(2)),
                    },
                    Mode::Passive,
                ],
                elements: vec![
                    element(Mode::Active {
                        index: 0,
                        offset: ConstExpr::Value(Value::I32(0)),
                    }),
                    element(Mode::Passive),
                    element(Mode::Declarative),
                ],
            }));
            let mut vm = VM::new();
            if linked {
                vm.run_linked(chunks, vec![]).unwrap();
            } else {
                vm.run(chunks).unwrap();
            }
            vm.memory
                .with_buffer_mut(|bytes| assert_eq!(&bytes[2..5], &[42, 43, 44]));
            assert!(vm.data_segments[0].is_empty());
            assert_eq!(vm.data_segments[0].capacity(), 0);
            assert_eq!(vm.data_segments[1], [81]);
            assert!(vm.dropped_data.contains(&0));
            assert!(!vm.dropped_data.contains(&1));
            assert!(matches!(vm.wasm_tables[0][0], Value::Object(_)));
            assert!(vm.elem_segments[0].is_empty());
            assert_eq!(vm.elem_segments[0].capacity(), 0);
            assert_eq!(vm.elem_segments[1].len(), 1);
            assert!(vm.elem_segments[2].is_empty());
            assert_eq!(vm.elem_segments[2].capacity(), 0);
            assert!(vm.dropped_elems.contains(&0));
            assert!(!vm.dropped_elems.contains(&1));
            assert!(vm.dropped_elems.contains(&2));
        }
    }
}

impl Snapshot {
    pub(super) fn take(vm: &mut VM) -> Self {
        Self {
            memory: std::mem::take(&mut vm.memory),
            extra_memories: std::mem::take(&mut vm.extra_memories),
            extra_memory_max_pages: std::mem::take(&mut vm.extra_memory_max_pages),
            memory_is_64: std::mem::take(&mut vm.memory_is_64),
            memory0_is_64: std::mem::take(&mut vm.memory0_is_64),
            active_memory: std::mem::take(&mut vm.active_memory),
            tables: std::mem::take(&mut vm.wasm_tables),
            table_maxes: std::mem::take(&mut vm.wasm_table_maxes),
            table_is_64: std::mem::take(&mut vm.table_is_64),
            table0_is_64: std::mem::take(&mut vm.table0_is_64),
            data: std::mem::take(&mut vm.data_segments),
            elements: std::mem::take(&mut vm.elem_segments),
            dropped_data: std::mem::take(&mut vm.dropped_data),
            dropped_elements: std::mem::take(&mut vm.dropped_elems),
        }
    }

    pub(super) fn restore(self, vm: &mut VM) {
        vm.memory = self.memory;
        vm.extra_memories = self.extra_memories;
        vm.extra_memory_max_pages = self.extra_memory_max_pages;
        vm.memory_is_64 = self.memory_is_64;
        vm.memory0_is_64 = self.memory0_is_64;
        vm.active_memory = self.active_memory;
        vm.wasm_tables = self.tables;
        vm.wasm_table_maxes = self.table_maxes;
        vm.table_is_64 = self.table_is_64;
        vm.table0_is_64 = self.table0_is_64;
        vm.data_segments = self.data;
        vm.elem_segments = self.elements;
        vm.dropped_data = self.dropped_data;
        vm.dropped_elems = self.dropped_elements;
    }
}

impl VM {
    pub(super) fn initialize_module_storage(&mut self, script: usize) -> Result<(), VMError> {
        let memories = self.chunks[script].memory_min_pages.clone();
        let maxima = self.chunks[script].memory_max_pages.clone();
        self.memory_is_64 = self.chunks[script].memory_is_64.clone();
        self.memory0_is_64 = self.memory_is_64.first().copied().unwrap_or(false);
        self.instantiate_declared_memories(&memories, &maxima)?;
        self.table_is_64 = self.chunks[script].table_is_64.clone();
        self.table0_is_64 = self.table_is_64.first().copied().unwrap_or(false);
        let tables = self.chunks[script].table_min_sizes.clone();
        self.wasm_table_maxes = self.chunks[script]
            .table_max_sizes
            .iter()
            .map(|maximum| maximum.map(|size| size as usize))
            .collect();
        self.instantiate_declared_tables(&tables)?;
        let data = self.chunks[script].data_segments.clone();
        if self.chunks[script].native_segments.is_some() || !data.is_empty() {
            self.data_segments = data;
            self.dropped_data.clear();
        }
        Ok(())
    }

    pub(super) fn initialize_module_segments(&mut self, script: usize) -> Result<(), VMError> {
        if let Some(segments) = self.chunks[script].native_segments.clone() {
            return self.initialize_native_segments(script, &segments);
        }
        // References and expression elements require installed types/globals.
        let elements = self.chunks[script].elem_segments.clone();
        if !elements.is_empty() {
            let base = if self.chunks[script].name == "<script>" && self.chunks.len() > script + 1 {
                script + 1
            } else {
                script
            };
            self.elem_segments = elements
                .into_iter()
                .map(|segment| {
                    segment
                        .into_iter()
                        .map(|value| match value {
                            Value::I32(index) if index >= 0 => {
                                let chunk = base + index as usize;
                                if chunk < self.chunks.len() {
                                    self.make_funcref(chunk)
                                } else {
                                    Value::Null
                                }
                            }
                            other => other,
                        })
                        .collect()
                })
                .collect();
            self.dropped_elems.clear();
        }
        let items = self.chunks[script].passive_elem_items.clone();
        let type_base = self.chunk_type_base.get(script).copied().unwrap_or(0);
        for (index, expressions) in items.iter().enumerate() {
            let values = expressions
                .iter()
                .map(|expression| self.eval_const_expr_with_context(expression, type_base, script))
                .collect::<Result<_, _>>()?;
            self.set_elem_segment(index, values);
        }
        for init in self.chunks[script].active_elem_segments.clone() {
            let values = self
                .elem_segments
                .get(init.elem_index as usize)
                .ok_or_else(|| VMError::new("active element segment payload missing"))?;
            let offset = usize::try_from(init.offset)
                .map_err(|_| VMError::new("active element segment offset out of range"))?;
            let table = self
                .wasm_tables
                .get_mut(init.table_index as usize)
                .ok_or_else(|| VMError::new("active element segment table missing"))?;
            let end = offset
                .checked_add(values.len())
                .filter(|end| *end <= table.len())
                .ok_or_else(|| VMError::new("trap: active element segment out of bounds"))?;
            table[offset..end].clone_from_slice(values);
        }
        for init in self.chunks[script].active_data_segments.clone() {
            let bytes = self
                .data_segments
                .get(init.data_index as usize)
                .ok_or_else(|| VMError::new("active data segment payload missing"))?;
            let offset = usize::try_from(init.offset)
                .map_err(|_| VMError::new("active data segment offset out of range"))?;
            let memory = if init.memory_index == 0 {
                &self.memory
            } else {
                self.extra_memories
                    .get(init.memory_index as usize - 1)
                    .ok_or_else(|| VMError::new("active data segment memory missing"))?
            };
            memory.with_buffer_mut(|memory| {
                let end = Self::checked_memory_end(offset, bytes.len(), memory.len())?;
                memory[offset..end].copy_from_slice(bytes);
                Ok::<_, VMError>(())
            })?;
        }
        Ok(())
    }

    fn initialize_native_segments(
        &mut self,
        script: usize,
        segments: &crate::chunk::NativeSegments,
    ) -> Result<(), VMError> {
        use crate::chunk::NativeSegmentMode;
        let type_base = self.chunk_type_base.get(script).copied().unwrap_or(0);
        self.elem_segments.clear();
        self.dropped_elems.clear();
        for (index, segment) in segments.elements.iter().enumerate() {
            let values = segment
                .items
                .iter()
                .map(|item| self.eval_const_expr_with_context(item, type_base, script))
                .collect::<Result<Vec<_>, _>>()?;
            self.elem_segments.push(values);
            if matches!(segment.mode, NativeSegmentMode::Declarative) {
                self.elem_segments[index] = Vec::new();
                self.dropped_elems.insert(index as u32);
            }
        }
        // Check every active range before applying any initializer.
        let mut element_offsets = Vec::with_capacity(segments.elements.len());
        for (index, segment) in segments.elements.iter().enumerate() {
            if let NativeSegmentMode::Active {
                index: table,
                offset,
            } = &segment.mode
            {
                let offset = self.native_segment_offset(offset, type_base, script)?;
                let table_len = self
                    .wasm_tables
                    .get(*table as usize)
                    .ok_or_else(|| VMError::new("active element segment table missing"))?
                    .len();
                let end = offset
                    .checked_add(self.elem_segments[index].len())
                    .filter(|end| *end <= table_len)
                    .ok_or_else(|| VMError::new("trap: active element segment out of bounds"))?;
                element_offsets.push((index, *table as usize, offset, end));
            }
        }
        let mut data_offsets = Vec::with_capacity(segments.data_modes.len());
        for (index, mode) in segments.data_modes.iter().enumerate() {
            if let NativeSegmentMode::Active {
                index: memory,
                offset,
            } = mode
            {
                let offset = self.native_segment_offset(offset, type_base, script)?;
                let length = self
                    .data_segments
                    .get(index)
                    .ok_or_else(|| VMError::new("active data segment payload missing"))?
                    .len();
                let memory = if *memory == 0 {
                    self.memory.clone()
                } else {
                    self.extra_memories
                        .get(*memory as usize - 1)
                        .ok_or_else(|| VMError::new("active data segment memory missing"))?
                        .clone()
                };
                let end = offset
                    .checked_add(length)
                    .filter(|end| *end <= memory.len())
                    .ok_or_else(|| VMError::new("trap: active data segment out of bounds"))?;
                data_offsets.push((index, memory, offset, end));
            }
        }
        for (index, table, offset, end) in element_offsets {
            self.wasm_tables[table][offset..end].clone_from_slice(&self.elem_segments[index]);
            self.elem_segments[index] = Vec::new();
            self.dropped_elems.insert(index as u32);
        }
        for (index, memory, offset, end) in data_offsets {
            memory.with_buffer_mut(|buffer| {
                buffer[offset..end].copy_from_slice(&self.data_segments[index])
            });
            self.data_segments[index] = Vec::new();
            self.dropped_data.insert(index as u32);
        }
        Ok(())
    }

    fn native_segment_offset(
        &mut self,
        expression: &crate::chunk::ConstExpr,
        types: usize,
        script: usize,
    ) -> Result<usize, VMError> {
        let offset = match self.eval_const_expr_with_context(expression, types, script)? {
            Value::I32(value) => u64::from(value as u32),
            Value::I64(value) => value as u64,
            _ => return Err(VMError::new("segment offset is not an integer")),
        };
        usize::try_from(offset).map_err(|_| VMError::new("segment offset out of range"))
    }
}
