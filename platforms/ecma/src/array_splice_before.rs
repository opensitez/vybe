// Frozen pre-change splice provider; benchmark feature only.
use super::*;
pub(super) fn splice(ctx: &mut HostContext, args: &[Value]) -> Value {
    let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
    // §23.1.3.31: when deleteCount is OMITTED, everything from
    // start to the end is removed (explicit undefined means 0).
    let del = match args.get(2) {
        None => usize::MAX,
        Some(v) => v.as_i32().max(0) as usize,
    };
    let items = args.get(3..).unwrap_or(&[]);
    let mut deleted = Vec::new();
    let mut deleted_holes = BTreeSet::new();
    if let Some(arr) = array_of(args, 0) {
        if is_frozen(&arr) {
            return throw_type_error(ctx, "Cannot modify frozen array");
        }
        let mut o = arr.lock().unwrap();
        if let ObjectKind::Array(ref v) = o.kind {
            let len = v.len();
            let idx = if start < 0 {
                ((len as i32) + start).max(0) as usize
            } else {
                (start as usize).min(len)
            };
            let end = idx.saturating_add(del).min(len);
            let delete_count = end.saturating_sub(idx);
            let insert_count = items.len();
            deleted.reserve(delete_count);
            let old_holes = hole_indices(&o);
            for offset in 0..delete_count {
                if old_holes.contains(&(idx + offset)) {
                    deleted_holes.insert(offset);
                }
            }
            if let ObjectKind::Array(ref mut v) = o.kind {
                if delete_count <= 1 && insert_count <= 1 {
                    // Keep the small queue operation on its original path.
                    for _ in idx..end {
                        deleted.push(v.remove(idx));
                    }
                    for (i, val) in items.iter().cloned().enumerate() {
                        v.insert(idx + i, val);
                    }
                } else {
                    // Bulk replacement moves the tail once. Value cloning
                    // invokes no guest callbacks; metadata stays below.
                    deleted.extend(v.splice(idx..end, items.iter().cloned()));
                }
            }
            let shift = insert_count as isize - delete_count as isize;
            let remapped: BTreeSet<usize> = old_holes
                .into_iter()
                .filter_map(|hole| {
                    if hole < idx {
                        Some(hole)
                    } else if hole < end {
                        None
                    } else {
                        Some((hole as isize + shift) as usize)
                    }
                })
                .collect();
            store_hole_indices(&mut o, &remapped);
        }
        sync_length(&mut o);
    }
    let removed = make_array(deleted);
    if let Value::Object(obj) = &removed {
        let mut guard = obj.lock().unwrap();
        store_hole_indices(&mut guard, &deleted_holes);
    }
    removed
}
