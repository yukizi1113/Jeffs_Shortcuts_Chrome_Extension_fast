// Y-U0036 Windows-like keyboard settings for macOS
// VENDOR_ID と PRODUCT_ID を EventViewer の実際の値に変更してください。

function main() {
    var VENDOR_ID = 1133;   // ← 実際の vendor_id に変更
    var PRODUCT_ID = 49995;  // ← 実際の product_id に変更

    function deviceCondition() {
        return {
            type: 'device_if',
            identifiers: [
                {
                    vendor_id: VENDOR_ID,
                    product_id: PRODUCT_ID,
                    is_keyboard: true
                }
            ]
        };
    }

    function browserCondition() {
        return {
            type: 'frontmost_application_if',
            bundle_identifiers: [
                '^com\\.google\\.Chrome.*$',
                '^com\\.microsoft\\.edgemac.*$'
            ]
        };
    }

function excelUnlessCondition() {
    return {
        type: 'frontmost_application_unless',
        bundle_identifiers: [
            '^com\\.microsoft\\.Excel$'
        ]
    };
}

function excelCondition() {
    return {
        type: 'frontmost_application_if',
        bundle_identifiers: [
            '^com\\.microsoft\\.Excel$'
        ]
    };
}    
    
    function macTextAppCondition() {
        return {
            type: 'frontmost_application_if',
            bundle_identifiers: [
                '^com\\.apple\\.TextEdit$',
                '^com\\.apple\\.Notes$'
            ]
        };
    }


function jeffBrowserCondition() {
    return {
        type: 'frontmost_application_if',
        bundle_identifiers: [
            '^com\\.google\\.Chrome.*$'
        ]
    };
}

    function finderCondition() {
        return {
            type: 'frontmost_application_if',
            bundle_identifiers: [
                '^com\\.apple\\.finder$'
            ]
        };
    }

    var manipulators = [];

//
// ============================================================
// Excel for Mac - Windows QAT Alt+1 / Alt+2 / Alt+3
// Y-U0036 + Microsoft Excel のときだけ有効
// ============================================================
//

// Alt + 1 → Ctrl + Shift + Y
// → QAT_TracePrecedents
manipulators.unshift({
    type: 'basic',
    from: {
        key_code: '1',
        modifiers: {
            mandatory: ['option'],
            optional: ['caps_lock']
        }
    },
    to: [
        {
            key_code: 'y',
            modifiers: ['left_control', 'left_shift']
        }
    ],
    conditions: [
        deviceCondition(),
        excelCondition()
    ]
});

// Alt + 2 → Ctrl + Shift + G
// → QAT_TraceDependents
manipulators.unshift({
    type: 'basic',
    from: {
        key_code: '2',
        modifiers: {
            mandatory: ['option'],
            optional: ['caps_lock']
        }
    },
    to: [
        {
            key_code: 'g',
            modifiers: ['left_control', 'left_shift']
        }
    ],
    conditions: [
        deviceCondition(),
        excelCondition()
    ]
});

// Alt + 3 → Ctrl + Shift + X
// → QAT_ClearArrows
manipulators.unshift({
    type: 'basic',
    from: {
        key_code: '3',
        modifiers: {
            mandatory: ['option'],
            optional: ['caps_lock']
        }
    },
    to: [
        {
            key_code: 'x',
            modifiers: ['left_control', 'left_shift']
        }
    ],
    conditions: [
        deviceCondition(),
        excelCondition()
    ]
});


//
// ============================================================
// Excel for Mac：Windows風 Ctrl + Arrow
// ============================================================
// Y-U0036 + Microsoft Excel のときだけ有効
//

function addExcelCtrlArrow(key) {

    // Ctrl + Arrow
    // → Command + Arrow
    manipulators.push({
        type: 'basic',
        from: {
            key_code: key,
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: key,
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            excelCondition()
        ]
    });

    // Ctrl + Shift + Arrow
    // → Command + Shift + Arrow
    // Windows同様「データ端まで範囲選択」
    manipulators.push({
        type: 'basic',
        from: {
            key_code: key,
            modifiers: {
                mandatory: ['control', 'shift'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: key,
                modifiers: ['left_command', 'left_shift']
            }
        ],
        conditions: [
            deviceCondition(),
            excelCondition()
        ]
    });
}

addExcelCtrlArrow('up_arrow');
addExcelCtrlArrow('down_arrow');
addExcelCtrlArrow('left_arrow');
addExcelCtrlArrow('right_arrow');


// ============================================================
// Excel for Mac - Jeff Alt shortcuts
// ============================================================
// Y-U0036 + Microsoft Excel のときだけ有効
//

//
// Ctrl + Alt + W
// → Jeff: DefaultSquareWidth
//
manipulators.unshift({
    type: 'basic',
    from: {
        key_code: 'w',
        modifiers: {
            mandatory: ['control', 'option'],
            optional: ['caps_lock']
        }
    },
    to: [
        {
            shell_command:
                "/usr/bin/osascript -e 'tell application \"Microsoft Excel\" to run VB macro \"DefaultSquareWidth\"'"
        }
    ],
    conditions: [
        deviceCondition(),
        excelCondition()
    ]
});


//
// Ctrl + Shift + Alt + F
// → Jeff: SetToDefaultFont
//
manipulators.unshift({
    type: 'basic',
    from: {
        key_code: 'f',
        modifiers: {
            mandatory: ['control', 'shift', 'option'],
            optional: ['caps_lock']
        }
    },
    to: [
        {
            shell_command:
                "/usr/bin/osascript -e 'tell application \"Microsoft Excel\" to run VB macro \"SetToDefaultFont\"'"
        }
    ],
    conditions: [
        deviceCondition(),
        excelCondition()
    ]
});

//
// macOS テキスト系アプリ：
// Windows Ctrl + Shift + V
// → macOS「ペーストしてスタイルを合わせる」
// Command + Option + Shift + V
//
manipulators.unshift({
    type: 'basic',
    from: {
        key_code: 'v',
        modifiers: {
            mandatory: ['control', 'shift'],
            optional: ['caps_lock']
        }
    },
    to: [
        {
            key_code: 'v',
            modifiers: [
                'left_command',
                'left_option',
                'left_shift'
            ]
        }
    ],
    conditions: [
        deviceCondition(),
        macTextAppCondition()
    ]
});

    //
    // 1. Alt + Tab → Windowsと同じアプリ切り替え
    //
    // 単純に Option+Tab → Command+Tab とすると、
    // Altを押し続けた連続切替で問題が出るため、
    // Karabiner公式推奨方式を使用。
    //
    ['left_option', 'right_option'].forEach(function (optionKey) {
        var commandKey =
            optionKey === 'left_option' ? 'left_command' : 'right_command';

        manipulators.push({
            type: 'basic',
            from: {
                key_code: optionKey,
                modifiers: { optional: ['any'] }
            },
            to: [
                { key_code: optionKey }
            ],
            to_if_other_key_pressed: [
                {
                    other_keys: [
                        {
                            key_code: 'tab',
                            modifiers: { optional: ['any'] }
                        }
                    ],
                    to: [
                        { key_code: commandKey }
                    ]
                }
            ],
            conditions: [
                deviceCondition()
            ]
        });
    });

    //
    // 2. Windowsの基本的な Ctrl ショートカット
    //
    // Ctrl+C → Command+C
    // Ctrl+V → Command+V
    // Ctrl+X → Command+X
    // Ctrl+Z → Command+Z
    // Ctrl+A → Command+A
    // Ctrl+S → Command+S
    // Ctrl+F → Command+F
    // Ctrl+P → Command+P
    // Ctrl+N → Command+N
    // Ctrl+O → Command+O
    // Ctrl+W → Command+W
    // Ctrl+T → Command+T
    // Ctrl+L → Command+L
    // Ctrl+R → Command+R
    // Ctrl+B/I/U/K/D/G など
    //
    // Shiftはそのまま維持されるため、
    // Ctrl+Shift+T → Command+Shift+T なども動作。
    //
    var ctrlToCommandKeys = [
        'a', 'b', 'c', 'd',
        'f', 'g', 'i', 'k', 'l',
        'n', 'o', 'p', 'r', 's', 't', 'u',
        'v', 'w', 'x', 'z',
        '0', '1', '2', '3', '4',
        '5', '6', '7', '8', '9',
        'equal_sign', 'hyphen'
    ];

    ctrlToCommandKeys.forEach(function (key) {
        manipulators.push({
            type: 'basic',
            from: {
                key_code: key,
                modifiers: {
                    mandatory: ['control'],
                    optional: ['shift', 'caps_lock']
                }
            },
            to: [
                {
                    key_code: key,
                    modifiers: ['left_command']
                }
            ],
            conditions: [
                deviceCondition(),
                excelUnlessCondition()
            ]
        });
    });

    //
    // 3. Chrome / EdgeだけWindowsショートカットを補正
    //
    // Windows Ctrl+H = 履歴
    // Mac Chrome Command+Y = 履歴
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'h',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'y',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });

    // Ctrl+J → ダウンロード
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'j',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'j',
                modifiers: ['left_command', 'left_shift']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });

    // Ctrl+U → ページのHTMLソース
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'u',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'u',
                modifiers: ['left_command', 'left_option']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });

    // Ctrl+Shift+I → DevTools
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'i',
            modifiers: {
                mandatory: ['control', 'shift'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'i',
                modifiers: ['left_command', 'left_option']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });

    // Ctrl+Shift+C → 要素を検証
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'c',
            modifiers: {
                mandatory: ['control', 'shift'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'c',
                modifiers: ['left_command', 'left_option']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });

    //
    // 4. Alt + F4 → 現在のウインドウを閉じる
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'f4',
            modifiers: {
                mandatory: ['option'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'w',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 5. Ctrl + F4 → 現在のタブ/ウインドウを閉じる
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'f4',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'w',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 6. Ctrl + Shift + Esc → 強制終了画面
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'escape',
            modifiers: {
                mandatory: ['control', 'shift'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'escape',
                modifiers: ['left_command', 'left_option']
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 7. 半角/全角 → 日本語 ⇔ 英数字
    //
    // 日本語入力中なら英数へ
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'grave_accent_and_tilde'
        },
        to: [
            {
                key_code: 'japanese_eisuu'
            }
        ],
        conditions: [
            deviceCondition(),
            {
                type: 'input_source_if',
                input_sources: [
                    { language: '^ja$' }
                ]
            }
        ]
    });

    // 英数字入力中なら日本語へ
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'grave_accent_and_tilde'
        },
        to: [
            {
                key_code: 'japanese_kana'
            }
        ],
        conditions: [
            deviceCondition(),
            {
                type: 'input_source_unless',
                input_sources: [
                    { language: '^ja$' }
                ]
            }
        ]
    });


    //
    // 7-2. 「カタカナ ひらがな ローマ字」→ 必ず日本語入力
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'japanese_pc_katakana',
            modifiers: {
                optional: ['any']
            }
        },
        to: [
            {
                key_code: 'japanese_kana'
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 7-3. Finder：Windows風ナビゲーション／切り取り＆移動
    //

    //
    // Alt + ↑ → 1つ上のフォルダへ
    // Windows: Alt+Up
    // macOS Finder: Command+Up
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'up_arrow',
            modifiers: {
                mandatory: ['option'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'up_arrow',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition()
        ]
    });

    //
    // Ctrl + X → 「切り取り予定」としてコピー
    //
    // macOS FinderにはファイルのCommand+Xがないため、
    // Command+Cを実行して変数に「次の貼り付けは移動」と記録する。
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'x',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'c',
                modifiers: ['left_command']
            },
            {
                set_variable: {
                    name: 'finder_cut_pending',
                    value: true
                }
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition()
        ]
    });

    //
    // Ctrl + C → 通常コピー
    // 以前のCtrl+X状態も解除する。
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'c',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'c',
                modifiers: ['left_command']
            },
            {
                set_variable: {
                    name: 'finder_cut_pending',
                    value: false
                }
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition()
        ]
    });

    //
    // Ctrl + V（Ctrl+X後）
    // → Finderの「ここに移動」= Command+Option+V
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'v',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'v',
                modifiers: [
                    'left_command',
                    'left_option'
                ]
            },
            {
                set_variable: {
                    name: 'finder_cut_pending',
                    value: false
                }
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition(),
            {
                type: 'variable_if',
                name: 'finder_cut_pending',
                value: true
            }
        ]
    });

    //
    // Ctrl + V（通常コピー後）
    // → Command+V
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'v',
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'v',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition(),
            {
                type: 'variable_unless',
                name: 'finder_cut_pending',
                value: true
            }
        ]
    });

    //
    // 7-4. Finder：Enter → 選択中のファイル/フォルダを開く
    //
    // Windows風:
    // Enter → 開く
    //
    // macOS Finder:
    // Command + O → 開く
    //
    // 名前変更中や検索欄など、テキスト入力中のEnterは変換しない。
    //
    manipulators.unshift({
        type: 'basic',
        from: {
            key_code: 'return_or_enter',
            modifiers: {
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: 'o',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition(),
            {
                type: 'expression_unless',
                expression: "accessibility.focused_ui_element.role_string like 'AXText*'"
            }
        ]
    });

    //
    // 8. Finder：F2 → 名前の変更
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'f2'
        },
        to: [
            {
                key_code: 'return_or_enter'
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition()
        ]
    });

    //
    // 9. Finder：Delete → ゴミ箱
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'delete_forward'
        },
        to: [
            {
                key_code: 'delete_or_backspace',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition()
        ]
    });

    //
    // 10. Finder：Shift + Delete → 即時削除
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'delete_forward',
            modifiers: {
                mandatory: ['shift']
            }
        },
        to: [
            {
                key_code: 'delete_or_backspace',
                modifiers: ['left_command', 'left_option']
            }
        ],
        conditions: [
            deviceCondition(),
            finderCondition()
        ]
    });

    //
    // 11. Windowsキー + L → 画面ロック
    //
    // WindowsキーはMacでは Command と認識される
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'l',
            modifiers: {
                mandatory: ['command']
            }
        },
        to: [
            {
                key_code: 'q',
                modifiers: ['left_control', 'left_command']
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 12. Windowsキー + E → Finder
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'e',
            modifiers: {
                mandatory: ['command']
            }
        },
        to: [
            {
                software_function: {
                    open_application: {
                        bundle_identifier: 'com.apple.finder'
                    }
                }
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 13. Windowsキー + R → Spotlight
    // Windowsの「ファイル名を指定して実行」に近い用途
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'r',
            modifiers: {
                mandatory: ['command']
            }
        },
        to: [
            {
                key_code: 'spacebar',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 14. Windowsキー + Shift + S → 範囲スクリーンショット
    // 結果をクリップボードへ
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 's',
            modifiers: {
                mandatory: ['command', 'shift']
            }
        },
        to: [
            {
                key_code: '4',
                modifiers: [
                    'left_control',
                    'left_shift',
                    'left_command'
                ]
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 15. Print Screen → 画面全体をクリップボードへ
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'print_screen'
        },
        to: [
            {
                key_code: '3',
                modifiers: [
                    'left_control',
                    'left_shift',
                    'left_command'
                ]
            }
        ],
        conditions: [
            deviceCondition()
        ]
    });

    //
    // 16. Chrome / Edge：F5 → 更新
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'f5'
        },
        to: [
            {
                key_code: 'r',
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });

    //
    // 17. Chrome / Edge：F12 → DevTools
    //
    manipulators.push({
        type: 'basic',
        from: {
            key_code: 'f12'
        },
        to: [
            {
                key_code: 'i',
                modifiers: ['left_command', 'left_option']
            }
        ],
        conditions: [
            deviceCondition(),
            browserCondition()
        ]
    });


    //
    // ============================================================
    // JEFF'S GOOGLE SHEETS FAST MODE
    // ============================================================
    //
    // Ctrl+Option+Shift+F12 でON/OFF。
    //
    // ONのとき:
    // Jeff'sショートカットについてはWindows→Mac変換を行わず、
    // 元のControl/Option/ShiftのままChromeへ送る。
    //
    // Chrome拡張がGoogle Sheets上で直接受け取り、
    // Sheets APIを呼び出す。
    //

    function jeffModeCondition() {
        return {
            type: 'variable_if',
            name: 'jeff_sheets_mode',
            value: true
        };
    }

    function jeffModeOffCondition() {
        return {
            type: 'variable_unless',
            name: 'jeff_sheets_mode',
            value: true
        };
    }

    var jeffManipulators = [];

    //
    // Jeff mode OFF → ON
    //
    jeffManipulators.push({
        type: 'basic',
        from: {
            key_code: '0',
            modifiers: {
                mandatory: ['control', 'option', 'shift'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                set_variable: {
                    name: 'jeff_sheets_mode',
                    value: true
                }
            },
            {
                set_notification_message: {
                    id: 'jeff_sheets_mode',
                    text: "Jeff's Sheets mode: ON"
                }
            }
        ],
        conditions: [
            deviceCondition(),
            jeffModeOffCondition()
        ]
    });

    //
    // Jeff mode ON → OFF
    //
    jeffManipulators.push({
        type: 'basic',
        from: {
            key_code: '0',
            modifiers: {
                mandatory: ['control', 'option', 'shift'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                set_variable: {
                    name: 'jeff_sheets_mode',
                    value: false
                }
            },
            {
                set_notification_message: {
                    id: 'jeff_sheets_mode',
                    text: "Jeff's Sheets mode: OFF"
                }
            }
        ],
        conditions: [
            deviceCondition(),
            jeffModeCondition()
        ]
    });

    //
    // Jeffショートカットをそのまま通すためのヘルパー
    //
    function addJeffPassthrough(key, mandatoryModifiers) {

        var outputModifiers = [];

        mandatoryModifiers.forEach(function(mod) {
            if (mod === 'control') {
                outputModifiers.push('left_control');
            } else if (mod === 'option') {
                outputModifiers.push('left_option');
            } else if (mod === 'shift') {
                outputModifiers.push('left_shift');
            } else if (mod === 'command') {
                outputModifiers.push('left_command');
            }
        });

        jeffManipulators.push({
            type: 'basic',

            from: {
                key_code: key,
                modifiers: {
                    mandatory: mandatoryModifiers,
                    optional: ['caps_lock']
                }
            },

            to: [
                {
                    key_code: key,
                    modifiers: outputModifiers
                }
            ],

            
            conditions: [
                deviceCondition(),
                jeffBrowserCondition(),
                jeffModeCondition()
            ]
        });
    }

    //
    // Formatting / colors
    //
    addJeffPassthrough('n', ['control','shift']);
    addJeffPassthrough('c', ['control','shift']);
    addJeffPassthrough('1', ['control','shift']);
    addJeffPassthrough('5', ['control','shift']);

    addJeffPassthrough(
        'c',
        ['control','option','shift']
    );

    addJeffPassthrough(
        'v',
        ['control','option','shift']
    );

    addJeffPassthrough('w', ['control','shift']);
    addJeffPassthrough('m', ['control','shift']);

    addJeffPassthrough(
        'm',
        ['control','option','shift']
    );

    addJeffPassthrough('t', ['control','shift']);
    addJeffPassthrough('i', ['control','shift']);

    addJeffPassthrough(
        'f',
        ['control','option','shift']
    );

    addJeffPassthrough('b', ['control','shift']);

    addJeffPassthrough(
        'b',
        ['control','option','shift']
    );

    //
    // SuperFill / formulas
    //
    addJeffPassthrough('d', ['control','shift']);
    addJeffPassthrough('d', ['control','option']);

    addJeffPassthrough('z', ['control','shift']);

    addJeffPassthrough('r', ['control','shift']);
    addJeffPassthrough('r', ['control','option']);

    addJeffPassthrough('q', ['control','shift']);

    addJeffPassthrough(
        'f4',
        ['control','shift']
    );

    addJeffPassthrough(
        'f4',
        ['control','option']
    );

    addJeffPassthrough('e', ['control','shift']);

    addJeffPassthrough(
        'e',
        ['control','option','shift']
    );

//
// Excel / Google Sheets風ナビゲーション
//
// Windows:
// Ctrl + Arrow
//   → データ領域の端まで移動
//
// Mac Google Sheets:
// Command + Arrow
//
// Y-U0036でJeff mode ONかつChrome前面の場合のみ変換する。
//

function addExcelNavigation(key) {
    jeffManipulators.push({
        type: 'basic',
        from: {
            key_code: key,
            modifiers: {
                mandatory: ['control'],
                optional: ['caps_lock']
            }
        },
        to: [
            {
                key_code: key,
                modifiers: ['left_command']
            }
        ],
        conditions: [
            deviceCondition(),
            jeffBrowserCondition(),
            jeffModeCondition()
        ]
    });
}

addExcelNavigation('up_arrow');
addExcelNavigation('down_arrow');
addExcelNavigation('left_arrow');
addExcelNavigation('right_arrow');

    //
    // Dimensions
    //
    addJeffPassthrough(
        'w',
        ['control','option','shift']
    );

    addJeffPassthrough(
        'h',
        ['control','option','shift']
    );

    addJeffPassthrough(
        'right_arrow',
        ['control','option','shift']
    );

    addJeffPassthrough(
        'left_arrow',
        ['control','option','shift']
    );

    addJeffPassthrough(
        'down_arrow',
        ['control','option','shift']
    );

    addJeffPassthrough(
        'up_arrow',
        ['control','option','shift']
    );

    //
    // Alignment
    //
    addJeffPassthrough('w', ['control','option']);
    addJeffPassthrough('l', ['control','option']);
    addJeffPassthrough('h', ['control','option']);

    //
    // Decimal
    //
    addJeffPassthrough(
        'comma',
        ['control','shift']
    );

    addJeffPassthrough(
        'period',
        ['control','shift']
    );

    //
    // Save / options
    //
    addJeffPassthrough('s', ['control','shift']);

    addJeffPassthrough(
        'o',
        ['control','option','shift']
    );

    //
    // Zoom
    //
    addJeffPassthrough('k', ['control','shift']);
    addJeffPassthrough('j', ['control','shift']);
    addJeffPassthrough('h', ['control','shift']);

    //
    // F1 help toggle
    //
    addJeffPassthrough('1', ['control','option']);

    //
    // Disabled shortcut
    //
    addJeffPassthrough(
        'f4',
        ['control','option','shift']
    );

    //
    // JeffルールをWindows風変換よりも必ず先に評価する
    //
    manipulators = jeffManipulators.concat(manipulators);

    return {
        description: 'Y-U0036 - Windows-like shortcuts',
        manipulators: manipulators
    };
}

main();