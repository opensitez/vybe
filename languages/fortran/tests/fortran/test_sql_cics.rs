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
    sql_include_01,
    "program p
implicit none
EXEC SQL INCLUDE SQLCA END-EXEC
print *, 1
end program p"
);
c!(
    sql_select_02,
    "program p
implicit none
integer :: id
EXEC SQL SELECT 1 INTO :id END-EXEC
print *, id
end program p"
);
c!(
    sql_insert_03,
    "program p
implicit none
integer :: id
id = 1
EXEC SQL INSERT INTO T(ID) VALUES(:id) END-EXEC
end program p"
);
c!(
    sql_update_04,
    "program p
implicit none
integer :: id
id = 1
EXEC SQL UPDATE T SET ID = :id END-EXEC
end program p"
);
c!(
    sql_delete_05,
    "program p
implicit none
EXEC SQL DELETE FROM T WHERE ID = 1 END-EXEC
end program p"
);
c!(
    sql_commit_06,
    "program p
implicit none
EXEC SQL COMMIT END-EXEC
end program p"
);
c!(
    sql_rollback_07,
    "program p
implicit none
EXEC SQL ROLLBACK END-EXEC
end program p"
);
c!(
    sql_cursor_decl_08,
    "program p
implicit none
EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM T END-EXEC
end program p"
);
c!(
    sql_cursor_open_09,
    "program p
implicit none
EXEC SQL OPEN C1 END-EXEC
end program p"
);
c!(
    sql_cursor_fetch_10,
    "program p
implicit none
integer :: id
EXEC SQL FETCH C1 INTO :id END-EXEC
end program p"
);
c!(
    cics_link_11,
    "program p
implicit none
EXEC CICS LINK PROGRAM('SUB1') END-EXEC
end program p"
);
c!(
    cics_xctl_12,
    "program p
implicit none
EXEC CICS XCTL PROGRAM('SUB2') END-EXEC
end program p"
);
c!(
    cics_return_13,
    "program p
implicit none
EXEC CICS RETURN END-EXEC
end program p"
);
c!(
    cics_send_map_14,
    "program p
implicit none
EXEC CICS SEND MAP('M1') MAPSET('S1') END-EXEC
end program p"
);
c!(
    cics_recv_map_15,
    "program p
implicit none
EXEC CICS RECEIVE MAP('M1') MAPSET('S1') END-EXEC
end program p"
);
c!(
    cics_readq_ts_16,
    "program p
implicit none
EXEC CICS READQ TS QUEUE('Q1') END-EXEC
end program p"
);
c!(
    cics_writeq_ts_17,
    "program p
implicit none
EXEC CICS WRITEQ TS QUEUE('Q1') END-EXEC
end program p"
);
c!(
    cics_handle_cond_18,
    "program p
implicit none
EXEC CICS HANDLE CONDITION ERROR(10) END-EXEC
10 continue
end program p"
);
c!(
    sql_and_cics_19,
    "program p
implicit none
integer :: id
EXEC SQL SELECT 1 INTO :id END-EXEC
EXEC CICS RETURN END-EXEC
end program p"
);
c!(
    sql_dynamic_20,
    "program p
implicit none
character(len=40) :: stmt
EXEC SQL PREPARE S1 FROM :stmt END-EXEC
end program p"
);
c!(
    sql_cursor_close_21,
    "program p
implicit none
integer :: rc
EXEC SQL CLOSE C1 INTO :rc END-EXEC
end program p"
);
c!(
    sql_commit_work_release_22,
    "program p
implicit none
EXEC SQL COMMIT WORK RELEASE END-EXEC
end program p"
);
c!(
    sql_rollback_work_23,
    "program p
implicit none
EXEC SQL ROLLBACK WORK END-EXEC
end program p"
);
c!(
    sql_prepare_execute_24,
    "program p
implicit none
character(len=40) :: stmt
integer :: rc
stmt = 'INSERT INTO T(ID) VALUES(:id)'
EXEC SQL PREPARE S2 FROM :stmt END-EXEC
EXEC SQL EXECUTE S2 USING SQLCA END-EXEC
EXEC SQL GET DIAGNOSTICS :rc = RETURN_STATUS END-EXEC
end program p"
);
c!(
    sql_connect_25,
    "program p
implicit none
character(len=20) :: dsn
integer :: ret
EXEC SQL CONNECT TO :dsn USER 'SYS' USING 'pass' END-EXEC
EXEC SQL GET DIAGNOSTICS :ret = RETURN_STATUS END-EXEC
end program p"
);
c!(
    cics_syncpoint_26,
    "program p
implicit none
EXEC CICS SYNCPOINT
EXEC CICS SYNCPOINT ROLLBACK
EXEC CICS SYNCPOINT ROLLBACK(YES)
EXEC CICS SYNCPOINT KEEP
end program p"
);
c!(
    cics_handle_cond_multi_27,
    "program p
implicit none
EXEC CICS HANDLE CONDITION LTERM(10) INVALIDP(20) ENDPGM(30) END-EXEC
10 continue
20 continue
30 continue
end program p"
);
