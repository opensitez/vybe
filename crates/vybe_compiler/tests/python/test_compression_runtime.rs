//! gzip, zlib, bz2, lzma compression.

crate::runtime_case!(
    gzip_compress_decompress,
    "import gzip\nprint(gzip.decompress(gzip.compress(b'hello')))\n",
    "b'hello'"
);
crate::runtime_case!(
    zlib_compress_decompress,
    "import zlib\nprint(zlib.decompress(zlib.compress(b'hello')))\n",
    "b'hello'"
);
crate::runtime_case!(
    bz2_compress_decompress,
    "import bz2\nprint(bz2.decompress(bz2.compress(b'hello')))\n",
    "b'hello'"
);
crate::runtime_case!(
    lzma_compress_decompress,
    "import lzma\nprint(lzma.decompress(lzma.compress(b'hello')))\n",
    "b'hello'"
);
crate::runtime_case!(
    zlib_crc32,
    "import zlib\nprint(zlib.crc32(b'hello') != 0)\n",
    "True"
);
crate::runtime_case!(
    zlib_adler32,
    "import zlib\nprint(zlib.adler32(b'hello') != 0)\n",
    "True"
);
crate::runtime_case!(
    gzip_open_bytesio,
    "import gzip\nimport io\ndata = gzip.compress(b'abc')\nprint(gzip.open(io.BytesIO(data), 'rb').read())\n",
    "b'abc'"
);
crate::runtime_case!(
    zlib_compress_level,
    "import zlib\nprint(len(zlib.compress(b'x' * 100, level=9)) > 0)\n",
    "True"
);
crate::runtime_case!(
    zlib_decompressobj,
    "import zlib\nd = zlib.decompressobj()\nprint(d.decompress(zlib.compress(b'hi')))\n",
    "b'hi'"
);
crate::runtime_case!(
    gzip_bad_data,
    "import gzip\ntry:\n gzip.decompress(b'bad')\n print('ok')\nexcept Exception:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    zlib_wbits,
    "import zlib\nprint(len(zlib.compress(b'a', wbits=zlib.MAX_WBITS)) > 0)\n",
    "True"
);
crate::runtime_case!(
    bz2_compress_level,
    "import bz2\nprint(len(bz2.compress(b'abc', compresslevel=9)) > 0)\n",
    "True"
);
crate::runtime_case!(
    lzma_compress_preset,
    "import lzma\nprint(len(lzma.compress(b'abc', preset=6)) > 0)\n",
    "True"
);
crate::runtime_case!(
    zlib_compressobj,
    "import zlib\nc = zlib.compressobj()\nprint(c.compress(b'hi') + c.flush())\n",
    "b'hi'"
);
crate::runtime_case!(
    gzip_module_name,
    "import gzip\nprint(gzip.__name__)\n",
    "gzip"
);
crate::runtime_case!(
    zlib_module_name,
    "import zlib\nprint(zlib.__name__)\n",
    "zlib"
);
crate::runtime_case!(
    bz2_module_name,
    "import bz2\nprint(bz2.__name__)\n",
    "bz2"
);
crate::runtime_case!(
    lzma_module_name,
    "import lzma\nprint(lzma.__name__)\n",
    "lzma"
);
crate::runtime_case!(
    zlib_error,
    "import zlib\ntry:\n zlib.decompress(b'not zlib')\n print('ok')\nexcept zlib.error:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    gzip_empty,
    "import gzip\nprint(gzip.decompress(gzip.compress(b'')))\n",
    "b''"
);
crate::runtime_case!(
    zlib_empty,
    "import zlib\nprint(zlib.decompress(zlib.compress(b'')))\n",
    "b''"
);
crate::runtime_case!(
    bz2_empty,
    "import bz2\nprint(bz2.decompress(bz2.compress(b'')))\n",
    "b''"
);
crate::runtime_case!(
    lzma_empty,
    "import lzma\nprint(lzma.decompress(lzma.compress(b'')))\n",
    "b''"
);
crate::runtime_case!(
    zlib_max_wbits,
    "import zlib\nprint(zlib.MAX_WBITS > 0)\n",
    "True"
);
crate::runtime_case!(
    zlib_zlib_version,
    "import zlib\nprint(isinstance(zlib.ZLIB_VERSION, str))\n",
    "True"
);
crate::runtime_case!(
    gzip_mtime,
    "import gzip\nprint(hasattr(gzip, 'GzipFile'))\n",
    "True"
);
crate::runtime_case!(
    bz2_open,
    "import bz2\nprint(callable(bz2.open))\n",
    "True"
);
crate::runtime_case!(
    lzma_open,
    "import lzma\nprint(callable(lzma.open))\n",
    "True"
);
crate::runtime_case!(
    zlib_decompress_maxlen,
    "import zlib\nd = zlib.decompressobj()\ndata = zlib.compress(b'hello world')\nprint(len(d.decompress(data)) > 0)\n",
    "True"
);
crate::runtime_case!(
    gzip_compresslevel,
    "import gzip\nprint(len(gzip.compress(b'x', compresslevel=1)) > 0)\n",
    "True"
);
crate::runtime_case!(
    lzma_filters,
    "import lzma\nprint(hasattr(lzma, 'FILTER_LZMA2'))\n",
    "True"
);
crate::runtime_case!(
    bz2_decompress_incremental,
    "import bz2\nd = bz2.BZ2Decompressor()\ndata = bz2.compress(b'abc')\nprint(d.decompress(data))\n",
    "b'abc'"
);
crate::runtime_case!(
    zlib_crc32_same,
    "import zlib\nprint(zlib.crc32(b'hi') == zlib.crc32(b'hi'))\n",
    "True"
);
crate::runtime_case!(
    gzip_read_mode,
    "import gzip\nimport io\ndata = gzip.compress(b'test')\nf = gzip.GzipFile(fileobj=io.BytesIO(data), mode='rb')\nprint(f.read())\n",
    "b'test'"
);
crate::runtime_case!(
    zlib_windowbits,
    "import zlib\nprint(zlib.DEFLATED)\n",
    "8"
);
crate::runtime_case!(
    lzma_check,
    "import lzma\nprint(hasattr(lzma, 'CHECK_CRC32'))\n",
    "True"
);
crate::runtime_case!(
    gzip_bad_os,
    "import gzip\nprint(hasattr(gzip, 'FEXTRA'))\n",
    "True"
);
crate::runtime_case!(
    zlib_flush,
    "import zlib\nc = zlib.compressobj()\nout = c.compress(b'a') + c.flush(zlib.Z_FINISH)\nprint(len(out) > 0)\n",
    "True"
);
crate::runtime_case!(
    bz2_error,
    "import bz2\ntry:\n bz2.decompress(b'bad')\n print('ok')\nexcept Exception:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    lzma_format,
    "import lzma\nprint(hasattr(lzma, 'FORMAT_XZ'))\n",
    "True"
);
crate::runtime_case!(
    zlib_copy,
    "import zlib\nc1 = zlib.compressobj()\nc2 = c1.copy()\nprint(c1.compress(b'a') is not None)\n",
    "True"
);
crate::runtime_case!(
    gzip_isatty,
    "import gzip\nimport io\nf = gzip.GzipFile(fileobj=io.BytesIO(), mode='wb')\nprint(f.writable())\n",
    "True"
);
crate::runtime_case!(
    compression_roundtrip_unicode,
    "import zlib\ndata = 'hello é'.encode('utf-8')\nprint(zlib.decompress(zlib.compress(data)).decode('utf-8'))\n",
    "hello é"
);
crate::runtime_case!(
    gzip_header_crc,
    "import gzip\nprint(hasattr(gzip, 'FHCRC'))\n",
    "True"
);
crate::runtime_case!(
    lzma_preset_extreme,
    "import lzma\nprint(hasattr(lzma, 'PRESET_EXTREME'))\n",
    "True"
);

crate::compile_case!(zipfile_zip, "import zipfile\nzipfile.ZipFile\n");
crate::compile_case!(tarfile_open, "import tarfile\ntarfile.open\n");
crate::compile_case!(gzip_open_path, "import gzip\ngzip.open\n");
crate::compile_case!(lzma_lzmafile, "import lzma\nlzma.LZMAFile\n");
crate::compile_case!(zlib_compressobj_wbits, "import zlib\nzlib.compressobj(wbits=-zlib.MAX_WBITS)\n");
