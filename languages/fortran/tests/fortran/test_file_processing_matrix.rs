use super::helpers::compile_ok;

macro_rules! c {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

c!(
    fp_open_01,
    "program t\nopen(unit=10, file='a.dat')\nclose(10)\nend program t\n"
);
c!(
    fp_close_02,
    "program t\nopen(unit=10, file='a.dat')\nclose(unit=10)\nend program t\n"
);
c!(
    fp_read_03,
    "program t\ninteger :: x\nopen(unit=10, file='a.dat')\nread(10,*,end=100) x\n100 close(10)\nend program t\n"
);
c!(
    fp_write_04,
    "program t\ninteger :: x=1\nopen(unit=10, file='a.dat')\nwrite(10,*) x\nclose(10)\nend program t\n"
);
c!(
    fp_rewind_05,
    "program t\nopen(unit=10, file='a.dat')\nrewind(10)\nclose(10)\nend program t\n"
);
c!(
    fp_backspace_06,
    "program t\nopen(unit=10, file='a.dat')\nbackspace(10)\nclose(10)\nend program t\n"
);
c!(
    fp_endfile_07,
    "program t\nopen(unit=10, file='a.dat')\nendfile(10)\nclose(10)\nend program t\n"
);
c!(
    fp_flush_08,
    "program t\nopen(unit=10, file='a.dat')\nflush(10)\nclose(10)\nend program t\n"
);
c!(
    fp_inquire_09,
    "program t\nlogical :: opn\ninquire(unit=10, opened=opn)\nend program t\n"
);
c!(
    fp_stream_10,
    "program t\nopen(unit=10, file='a.dat', access='stream')\nclose(10)\nend program t\n"
);
c!(
    fp_direct_11,
    "program t\nopen(unit=10, file='a.dat', access='direct', recl=4)\nclose(10)\nend program t\n"
);
c!(
    fp_seq_12,
    "program t\nopen(unit=10, file='a.dat', access='sequential')\nclose(10)\nend program t\n"
);
c!(
    fp_formatted_13,
    "program t\nopen(unit=10, file='a.dat', form='formatted')\nclose(10)\nend program t\n"
);
c!(
    fp_unformatted_14,
    "program t\nopen(unit=10, file='a.dat', form='unformatted')\nclose(10)\nend program t\n"
);
c!(
    fp_status_old_15,
    "program t\nopen(unit=10, file='a.dat', status='old')\nclose(10)\nend program t\n"
);
c!(
    fp_status_new_16,
    "program t\nopen(unit=10, file='a.dat', status='new')\nclose(10)\nend program t\n"
);
c!(
    fp_status_replace_17,
    "program t\nopen(unit=10, file='a.dat', status='replace')\nclose(10)\nend program t\n"
);
c!(
    fp_action_read_18,
    "program t\nopen(unit=10, file='a.dat', action='read')\nclose(10)\nend program t\n"
);
c!(
    fp_action_write_19,
    "program t\nopen(unit=10, file='a.dat', action='write')\nclose(10)\nend program t\n"
);
c!(
    fp_action_rw_20,
    "program t\nopen(unit=10, file='a.dat', action='readwrite')\nclose(10)\nend program t\n"
);
c!(
    fp_position_append_21,
    "program t\nopen(unit=10, file='a.dat', position='append')\nclose(10)\nend program t\n"
);
c!(
    fp_position_rewind_22,
    "program t\nopen(unit=10, file='a.dat', position='rewind')\nclose(10)\nend program t\n"
);
c!(
    fp_rec_read_23,
    "program t\ninteger :: x\nopen(unit=10, file='a.dat', access='direct', recl=4)\nread(10,rec=1) x\nclose(10)\nend program t\n"
);
c!(
    fp_rec_write_24,
    "program t\ninteger :: x=1\nopen(unit=10, file='a.dat', access='direct', recl=4)\nwrite(10,rec=1) x\nclose(10)\nend program t\n"
);
c!(
    fp_iostat_25,
    "program t\ninteger :: ios\nopen(unit=10, file='a.dat', iostat=ios)\nclose(10)\nend program t\n"
);
c!(
    fp_err_label_26,
    "program t\nopen(unit=10, file='a.dat', err=100)\nclose(10)\n100 continue\nend program t\n"
);
c!(
    fp_namelist_27,
    "program t\ninteger :: x\nnamelist /grp/ x\nopen(unit=10, file='a.dat')\nwrite(10,nml=grp)\nclose(10)\nend program t\n"
);
c!(
    fp_pending_async_28,
    "program t\ninteger :: id\nopen(unit=10, file='a.dat', asynchronous='yes')\nwrite(10,asynchronous='yes',id=id) 1\nclose(10)\nend program t\n"
);
c!(
    fp_read_previous_shape_29,
    "program t\nopen(unit=10, file='a.dat')\nbackspace(10)\nclose(10)\nend program t\n"
);
c!(
    fp_inquire_file_30,
    "program t\nlogical :: ex\ninquire(file='a.dat', exist=ex)\nend program t\n"
);
