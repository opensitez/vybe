//! open/read/write, io.StringIO/BytesIO, text vs binary I/O.

crate::runtime_case!(
    io_stringio_write_read,
    "import io\ns = io.StringIO()\ns.write('hello')\ns.seek(0)\nprint(s.read())\n",
    "hello"
);
crate::runtime_case!(
    io_stringio_getvalue,
    "import io\ns = io.StringIO()\ns.write('abc')\nprint(s.getvalue())\n",
    "abc"
);
crate::runtime_case!(
    io_stringio_init_with_value,
    "import io\ns = io.StringIO('init')\nprint(s.read())\n",
    "init"
);
crate::runtime_case!(
    io_stringio_tell_seek,
    "import io\ns = io.StringIO('abcd')\ns.seek(2)\nprint(s.tell())\n",
    "2"
);
crate::runtime_case!(
    io_stringio_readline,
    "import io\ns = io.StringIO('a\\nb\\n')\nprint(s.readline().strip())\n",
    "a"
);
crate::runtime_case!(
    io_stringio_readlines,
    "import io\ns = io.StringIO('a\\nb\\n')\nprint(len(s.readlines()))\n",
    "2"
);
crate::runtime_case!(
    io_stringio_truncate,
    "import io\ns = io.StringIO('abcdef')\ns.truncate(3)\ns.seek(0)\nprint(s.read())\n",
    "abc"
);
crate::runtime_case!(
    io_bytesio_write_read,
    "import io\nb = io.BytesIO()\nb.write(b'hi')\nb.seek(0)\nprint(b.read())\n",
    "b'hi'"
);
crate::runtime_case!(
    io_bytesio_getvalue,
    "import io\nb = io.BytesIO()\nb.write(b'xyz')\nprint(b.getvalue())\n",
    "b'xyz'"
);
crate::runtime_case!(
    io_bytesio_getbuffer,
    "import io\nb = io.BytesIO(b'abc')\nprint(len(b.getbuffer()))\n",
    "3"
);
crate::runtime_case!(
    io_stringio_writelines,
    "import io\ns = io.StringIO()\ns.writelines(['a', 'b'])\nprint(s.getvalue())\n",
    "ab"
);
crate::runtime_case!(
    io_stringio_readable,
    "import io\ns = io.StringIO('x')\nprint(s.readable())\n",
    "True"
);
crate::runtime_case!(
    io_stringio_writable,
    "import io\ns = io.StringIO()\nprint(s.writable())\n",
    "True"
);
crate::runtime_case!(
    io_stringio_seekable,
    "import io\ns = io.StringIO()\nprint(s.seekable())\n",
    "True"
);
crate::runtime_case!(
    io_stringio_closed,
    "import io\ns = io.StringIO()\nprint(s.closed)\n",
    "False"
);
crate::runtime_case!(
    io_stringio_close,
    "import io\ns = io.StringIO()\ns.close()\nprint(s.closed)\n",
    "True"
);
crate::runtime_case!(
    io_stringio_flush,
    "import io\ns = io.StringIO()\ns.write('x')\ns.flush()\nprint(s.getvalue())\n",
    "x"
);
crate::runtime_case!(
    io_stringio_read_size,
    "import io\ns = io.StringIO('abcdef')\nprint(s.read(3))\n",
    "abc"
);
crate::runtime_case!(
    io_stringio_readinto_not_supported,
    "import io\ns = io.StringIO('ab')\nprint(hasattr(s, 'readinto'))\n",
    "True"
);
crate::runtime_case!(
    io_bytesio_read1,
    "import io\nb = io.BytesIO(b'hello')\nprint(b.read1(2))\n",
    "b'he'"
);
crate::runtime_case!(
    io_open_module_alias,
    "import io\nprint(callable(io.open))\n",
    "True"
);
crate::runtime_case!(
    io_textiowrapper_name,
    "import io\nprint(hasattr(io, 'TextIOWrapper'))\n",
    "True"
);
crate::runtime_case!(
    io_bufferedreader_name,
    "import io\nprint(hasattr(io, 'BufferedReader'))\n",
    "True"
);
crate::runtime_case!(
    io_bufferedwriter_name,
    "import io\nprint(hasattr(io, 'BufferedWriter'))\n",
    "True"
);
crate::runtime_case!(
    io_default_encoding,
    "import io\nprint(hasattr(io, 'DEFAULT_BUFFER_SIZE'))\n",
    "True"
);
crate::runtime_case!(
  print_to_stringio,
    "import io\nimport sys\nbuf = io.StringIO()\nprint('hi', file=buf)\nprint(buf.getvalue().strip())\n",
    "hi"
);
crate::runtime_case!(
    stringio_iteration,
    "import io\ns = io.StringIO('ab')\nprint(''.join(s))\n",
    "ab"
);
crate::runtime_case!(
    bytesio_iteration,
    "import io\nb = io.BytesIO(b'ab')\nprint(list(b))\n",
    "[97, 98]"
);
crate::runtime_case!(
    stringio_seek_end,
    "import io\ns = io.StringIO('abc')\ns.seek(0, 2)\nprint(s.tell())\n",
    "3"
);
crate::runtime_case!(
    stringio_seek_set,
    "import io\ns = io.StringIO('abc')\ns.seek(1)\nprint(s.read())\n",
    "bc"
);
crate::runtime_case!(
    stringio_newlines,
    "import io\ns = io.StringIO('a\\r\\nb')\nprint('\\n' in s.getvalue())\n",
    "True"
);
crate::runtime_case!(
    stringio_line_buffering_attr,
    "import io\ns = io.StringIO()\nprint(hasattr(s, 'line_buffering'))\n",
    "True"
);
crate::runtime_case!(
    bytesio_write_then_read,
    "import io\nb = io.BytesIO()\nb.write(b'\\x01\\x02')\nb.seek(0)\nprint(list(b.read()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    stringio_with_statement,
    "import io\nwith io.StringIO() as s:\n s.write('x')\n print(s.getvalue())\n",
    "x"
);
crate::runtime_case!(
    io_rawiobase_exists,
    "import io\nprint(hasattr(io, 'RawIOBase'))\n",
    "True"
);
crate::runtime_case!(
    io_iobase_exists,
    "import io\nprint(hasattr(io, 'IOBase'))\n",
    "True"
);
crate::runtime_case!(
    io_incrementalnewlinedecoder,
    "import io\nprint(hasattr(io, 'IncrementalNewlineDecoder'))\n",
    "True"
);
crate::runtime_case!(
    io_stringio_detach_none,
    "import io\ns = io.StringIO()\nprint(s.detach() is None)\n",
    "True"
);
crate::runtime_case!(
    io_bytesio_name_attr,
    "import io\nb = io.BytesIO()\nprint(hasattr(b, 'name'))\n",
    "True"
);
crate::runtime_case!(
    io_stringio_mode_attr,
    "import io\ns = io.StringIO()\nprint(hasattr(s, 'mode'))\n",
    "True"
);
crate::runtime_case!(
    io_read_write_roundtrip,
    "import io\ns = io.StringIO()\ns.write('data')\ns.seek(0)\nprint(len(s.read()))\n",
    "4"
);
crate::runtime_case!(
    io_multiline_read,
    "import io\ns = io.StringIO('line1\\nline2')\nprint(s.read().count('\\n'))\n",
    "1"
);
crate::runtime_case!(
    io_empty_read,
    "import io\ns = io.StringIO('')\nprint(repr(s.read()))\n",
    "''"
);
crate::runtime_case!(
    io_bytesio_empty,
    "import io\nb = io.BytesIO()\nprint(b.getvalue())\n",
    "b''"
);
crate::runtime_case!(
    io_stringio_unicode,
    "import io\ns = io.StringIO('é')\nprint(s.getvalue())\n",
    "é"
);

crate::compile_case!(open_read_write, "f = open(__file__)\ndata = f.read(10)\nf.close()\n");
crate::compile_case!(open_with_statement, "with open(__file__) as f:\n f.readline()\n");
crate::compile_case!(open_binary_mode, "with open(__file__, 'rb') as f:\n f.read(1)\n");
crate::compile_case!(open_write_mode, "import tempfile\nimport os\np = tempfile.mktemp()\nf = open(p, 'w')\nf.write('x')\nf.close()\nos.remove(p)\n");
crate::compile_case!(io_textiowrapper_reconfigure, "import io\ns = io.StringIO()\nhasattr(s, 'reconfigure')\n");
