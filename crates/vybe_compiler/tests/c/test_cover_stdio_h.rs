//! stdio.h file/stream APIs — one API per test.


c_run_cases! {
    fopen_fclose_mode => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f = fopen(\"/tmp/vybe_c_io.txt\", \"w\"); fprintf(f, \"x\"); fclose(f); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    fread_fwrite_binary => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_bin.dat\",\"wb\"); int v=7; fwrite(&v,sizeof(v),1,f); fclose(f); f=fopen(\"/tmp/vybe_c_bin.dat\",\"rb\"); int o=0; fread(&o,sizeof(o),1,f); fclose(f); printf(\"%d\\n\", o); return 0;",
        expect: ["7"]
    },
    fseek_ftell => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_seek.txt\",\"w+\"); fputs(\"abcd\",f); fseek(f,2,SEEK_SET); printf(\"%ld\\n\", ftell(f)); fclose(f); return 0;",
        expect: ["2"]
    },
    rewind_resets => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_rew.txt\",\"w+\"); fputs(\"z\",f); rewind(f); fputc('y',f); fclose(f); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    fflush_stdout => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"a\"); fflush(stdout); printf(\"\\n\"); return 0;",
        expect: ["a"]
    },
    feof_after_read => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_eof.txt\",\"w+\"); fputc('a',f); rewind(f); fgetc(f); printf(\"%d\\n\", feof(f)); fclose(f); return 0;",
        expect: ["1"]
    },
    ferror_clearerr => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_err.txt\",\"r\"); fgetc(f); int e=ferror(f); clearerr(f); printf(\"%d %d\\n\", e, ferror(f)); fclose(f); return 0;",
        expect: ["1 0"]
    },
    perror_emits_line => {
        includes: ["<stdio.h>", "<errno.h>"],
        decls: "",
        body: "errno=EINVAL; perror(\"vybe\"); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    fgets_reads_line => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_gets.txt\",\"w+\"); fputs(\"line\\n\",f); rewind(f); char buf[16]; fgets(buf,sizeof(buf),f); printf(\"%s\", buf); fclose(f); return 0;",
        expect: ["line"]
    },
    fputs_writes => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_puts2.txt\",\"w\"); fputs(\"ok\",f); fclose(f); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    fgetc_fputc => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_fc.txt\",\"w+\"); fputc('Q',f); rewind(f); printf(\"%c\\n\", fgetc(f)); fclose(f); return 0;",
        expect: ["Q"]
    },
    remove_file => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_rm.txt\",\"w\"); fclose(f); remove(\"/tmp/vybe_c_rm.txt\"); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    rename_file => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "FILE *f=fopen(\"/tmp/vybe_c_old.txt\",\"w\"); fclose(f); rename(\"/tmp/vybe_c_old.txt\",\"/tmp/vybe_c_new.txt\"); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
}

c_compile_cases! {
    setvbuf_compile => { includes: ["<stdio.h>"], decls: "", body: "char b[BUFSIZ]; setvbuf(stdout,b,_IOFBF,BUFSIZ); return 0;" },
    setbuf_compile => { includes: ["<stdio.h>"], decls: "", body: "char b[BUFSIZ]; setbuf(stdout,b); return 0;" },
    freopen_compile => { includes: ["<stdio.h>"], decls: "", body: "FILE *f = freopen(\"/tmp/vybe_c_fr.txt\",\"w\",stdout); if (f) fclose(f); return 0;" },
    tmpfile_compile => { includes: ["<stdio.h>"], decls: "", body: "FILE *f = tmpfile(); if (f) fclose(f); return 0;" },
    fgetpos_fsetpos_compile => { includes: ["<stdio.h>"], decls: "", body: "FILE *f=fopen(\"/tmp/vybe_c_pos.txt\",\"w+\"); fpos_t pos; fgetpos(f,&pos); fsetpos(f,&pos); fclose(f); return 0;" },
    vfprintf_compile => { includes: ["<stdio.h>", "<stdarg.h>"], decls: "void logit(const char *fmt, ...) { va_list ap; va_start(ap,fmt); vfprintf(stdout,fmt,ap); va_end(ap); }", body: "logit(\"%d\\n\",1); return 0;" },
    vprintf_compile => { includes: ["<stdio.h>", "<stdarg.h>"], decls: "", body: "return 0;" },
    vsprintf_compile => { includes: ["<stdio.h>", "<stdarg.h>"], decls: "", body: "char b[8]; return 0;" },
    vsnprintf_compile => { includes: ["<stdio.h>", "<stdarg.h>"], decls: "", body: "char b[8]; return 0;" },
}
