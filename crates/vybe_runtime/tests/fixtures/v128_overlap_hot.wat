;; Matched overlap loop: 128 copies of 32KiB.
;; Whole-program timing includes initialization and result validation.
(module
  (memory (export "memory") 1)
  (data (i32.const 0) "\0b\30\55\7a\9f\c4\e9\0e\33\58\7d\a2\c7\ec\11\36\5b\80\a5\ca\ef\14\39\5e\83\a8\cd\f2\17\3c\61\86\ab\d0\f5\1a\3f\64\89\ae\d3\f8\1d\42\67\8c\b1\d6\fb\20\45\6a\8f\b4\d9\fe\23\48\6d\92\b7\dc\01\26\4b\70\95\ba\df\04\29\4e\73\98\bd\e2\07\2c\51\76\9b\c0\e5\0a\2f\54\79\9e\c3\e8\0d\32\57\7c\a1\c6\eb\10\35\5a\7f\a4\c9\ee\13\38\5d\82\a7\cc\f1\16\3b\60\85\aa\cf\f4\19\3e\63\88\ad\d2\f7\1c\41\66\8b\b0\d5\fa\1f\44\69\8e\b3\d8\fd\22\47\6c\91\b6\db\00\25\4a\6f\94\b9\de\03\28\4d\72\97\bc\e1\06\2b\50\75\9a\bf\e4\09\2e\53\78\9d\c2\e7\0c\31\56\7b\a0\c5\ea\0f\34\59\7e\a3\c8\ed\12\37\5c\81\a6\cb\f0\15\3a\5f\84\a9\ce\f3\18\3d\62\87\ac\d1\f6\1b\40\65\8a\af\d4\f9\1e\43\68\8d\b2\d7\fc\21\46\6b\90\b5\da\ff\24\49\6e\93\b8\dd\02\27\4c\71\96\bb\e0\05\2a\4f\74\99\be\e3\08\2d\52\77\9c\c1\e6")
  (func (export "main") (export "_start") (result i32)
    (local $src i32) (local $dst i32) (local $len i32) (local $i i32)
    (local $last v128) (local $rep i32)
    (memory.copy (i32.const 256) (i32.const 0) (i32.const 256))
    (memory.copy (i32.const 512) (i32.const 0) (i32.const 512))
    (memory.copy (i32.const 1024) (i32.const 0) (i32.const 1024))
    (memory.copy (i32.const 2048) (i32.const 0) (i32.const 2048))
    (memory.copy (i32.const 4096) (i32.const 0) (i32.const 4096))
    (memory.copy (i32.const 8192) (i32.const 0) (i32.const 8192))
    (memory.copy (i32.const 16384) (i32.const 0) (i32.const 16384))
    (memory.copy (i32.const 32768) (i32.const 0) (i32.const 32768))
    i32.const 0 local.set $src
    i32.const 1 local.set $dst
    i32.const 32768 local.set $len
    block $outer_end
      loop $outer
        local.get $rep i32.const 128 i32.ge_u br_if $outer_end
        i32.const 0 local.set $i
        block $end
          loop $copy
            local.get $i local.get $len i32.ge_u br_if $end
            local.get $src local.get $i i32.add
            v128.load
            local.set $last
            local.get $dst local.get $i i32.add
            local.get $last
            v128.store
            ;; no extra instruction
            local.get $i i32.const 16 i32.add local.set $i
            br $copy
          end
        end
        local.get $rep i32.const 1 i32.add local.set $rep
        br $outer
      end
    end
    (if (i32.ne (i32.load (i32.const 32753)) (i32.const -2107820013))
      (then unreachable))
    i32.const 1)
)

