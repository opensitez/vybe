//! memoryview, buffer protocol, bytes-like slicing.

crate::runtime_case!(
    memoryview_from_bytes,
    "print(len(memoryview(b'abc')))\n",
    "3"
);
crate::runtime_case!(
    memoryview_index,
    "mv = memoryview(b'abc')\nprint(mv[1])\n",
    "98"
);
crate::runtime_case!(
    memoryview_slice,
    "mv = memoryview(b'abcdef')\nprint(bytes(mv[2:4]))\n",
    "b'cd'"
);
crate::runtime_case!(
    memoryview_tolist,
    "print(memoryview(b'ab').tolist())\n",
    "[97, 98]"
);
crate::runtime_case!(
    memoryview_tobytes,
    "print(memoryview(b'hi').tobytes())\n",
    "b'hi'"
);
crate::runtime_case!(
    memoryview_toreadonly,
    "mv = memoryview(b'abc')\nprint(mv.readonly)\n",
    "True"
);
crate::runtime_case!(
    memoryview_ndim,
    "print(memoryview(b'abc').ndim)\n",
    "1"
);
crate::runtime_case!(
    memoryview_itemsize,
    "print(memoryview(b'abc').itemsize)\n",
    "1"
);
crate::runtime_case!(
    memoryview_format,
    "print(memoryview(b'abc').format)\n",
    "B"
);
crate::runtime_case!(
    memoryview_strides,
    "print(memoryview(b'abc').strides)\n",
    "(1,)"
);
crate::runtime_case!(
    memoryview_shape,
    "print(memoryview(b'abc').shape)\n",
    "(3,)"
);
crate::runtime_case!(
    memoryview_cast,
    "mv = memoryview(b'abcd')\nprint(len(mv.cast('H')))\n",
    "2"
);
crate::runtime_case!(
    memoryview_release,
    "mv = memoryview(b'abc')\nmv.release()\nprint(str(mv).startswith('<released memory'))\n",
    "True"
);
crate::runtime_case!(
    memoryview_from_bytearray,
    "ba = bytearray(b'abc')\nmv = memoryview(ba)\nba[0] = ord('x')\nprint(bytes(mv[0:1]))\n",
    "b'x'"
);
crate::runtime_case!(
    memoryview_equality,
    "print(memoryview(b'ab') == memoryview(b'ab'))\n",
    "True"
);
crate::runtime_case!(
    memoryview_iteration,
    "print(list(memoryview(b'ab')))\n",
    "[97, 98]"
);
crate::runtime_case!(
    memoryview_hex,
    "print(memoryview(b'\\xff').hex())\n",
    "ff"
);
crate::runtime_case!(
    memoryview_obj,
    "mv = memoryview(b'abc')\nprint(mv.obj)\n",
    "b'abc'"
);
crate::runtime_case!(
    memoryview_nbytes,
    "print(memoryview(b'abc').nbytes)\n",
    "3"
);
crate::runtime_case!(
    memoryview_c_contiguous,
    "print(memoryview(b'abc').c_contiguous)\n",
    "True"
);
crate::runtime_case!(
    memoryview_f_contiguous,
    "print(memoryview(b'abc').f_contiguous)\n",
    "True"
);
crate::runtime_case!(
    memoryview_suboffsets,
    "print(memoryview(b'abc').suboffsets)\n",
    "None"
);
crate::runtime_case!(
    bytes_from_memoryview,
    "print(bytes(memoryview(b'xyz')))\n",
    "b'xyz'"
);
crate::runtime_case!(
    list_from_memoryview,
    "print(list(memoryview(b'\\x01\\x02')))\n",
    "[1, 2]"
);
crate::runtime_case!(
    memoryview_slice_assign_bytearray,
    "ba = bytearray(b'abcd')\nmv = memoryview(ba)\nmv[1:3] = b'XY'\nprint(ba)\n",
    "bytearray(b'aXYd')"
);
crate::runtime_case!(
    memoryview_negative_index,
    "print(memoryview(b'abc')[-1])\n",
    "99"
);
crate::runtime_case!(
    memoryview_step,
    "print(bytes(memoryview(b'abcdef')[::2]))\n",
    "b'ace'"
);
crate::runtime_case!(
    array_buffer_memoryview,
    "import array\na = array.array('i', [1, 2, 3])\nmv = memoryview(a)\nprint(mv.itemsize)\n",
    "4"
);
crate::runtime_case!(
    memoryview_overlapping_slice,
    "mv = memoryview(b'abcdef')\nprint(bytes(mv[1:4]))\n",
    "b'bcd'"
);
crate::runtime_case!(
    memoryview_bool,
    "print(bool(memoryview(b'')))\n",
    "False"
);
crate::runtime_case!(
    memoryview_bool_nonempty,
    "print(bool(memoryview(b'a')))\n",
    "True"
);
crate::runtime_case!(
    memoryview_len,
    "print(len(memoryview(b'hello')))\n",
    "5"
);
crate::runtime_case!(
    memoryview_repr,
    "print('memory' in repr(memoryview(b'a')))\n",
    "True"
);
crate::runtime_case!(
    bytearray_buffer_protocol,
    "ba = bytearray(b'abc')\nprint(len(memoryview(ba)))\n",
    "3"
);
crate::runtime_case!(
    memoryview_pickling_not_supported,
    "import pickle\ntry:\n pickle.dumps(memoryview(b'a'))\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    memoryview_hash_not_supported,
    "try:\n hash(memoryview(b'a'))\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    memoryview_multi_slice,
    "mv = memoryview(b'0123456789')\nprint(bytes(mv[2:8:2]))\n",
    "b'246'"
);
crate::runtime_case!(
    memoryview_from_bytes_readonly_assign,
    "mv = memoryview(b'abc')\ntry:\n mv[0] = 1\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    memoryview_tolist_empty,
    "print(memoryview(b'').tolist())\n",
    "[]"
);
crate::runtime_case!(
    memoryview_cast_format,
    "mv = memoryview(b'\\x00\\x01\\x00\\x02')\nprint(list(mv.cast('H')))\n",
    "[1, 2]"
);
crate::runtime_case!(
    buffer_bytes_conversion,
    "print(bytes(bytearray(b'abc')))\n",
    "b'abc'"
);
crate::runtime_case!(
    memoryview_in_contains,
    "mv = memoryview(b'abc')\nprint(98 in mv)\n",
    "True"
);
crate::runtime_case!(
    memoryview_count,
    "print(memoryview(b'aaa').count(ord('a')))\n",
    "3"
);
crate::runtime_case!(
    memoryview_find,
    "print(memoryview(b'abcabc').find(b'ca'))\n",
    "2"
);

crate::compile_case!(memoryview_from_array_multi, "import array\na = array.array('d', [1.0, 2.0])\nmemoryview(a)\n");
crate::compile_case!(memoryview_release_twice, "mv = memoryview(b'a')\nmv.release()\n");
crate::compile_case!(memoryview_cast_shape, "mv = memoryview(b'1234')\nmv.cast('I')\n");
crate::compile_case!(numpy_not_required, "mv = memoryview(b'abc')\n");
crate::compile_case!(memoryview_subclass, "class M(memoryview):\n pass\nM(b'a')\n");
