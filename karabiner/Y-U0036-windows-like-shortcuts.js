// Y-U0036 Windows-like keyboard settings for macOS
// Logitech Y-U0036: vendor_id=1133, product_id=49995

function main() {
    var VENDOR_ID = 1133;
    var PRODUCT_ID = 49995;

    function deviceCondition() {
        return {
            type: 'device_if',
            identifiers: [{ vendor_id: VENDOR_ID, product_id: PRODUCT_ID, is_keyboard: true }]
        };
    }

    function appIf(ids) {
        return { type: 'frontmost_application_if', bundle_identifiers: ids };
    }

    function appUnless(ids) {
        return { type: 'frontmost_application_unless', bundle_identifiers: ids };
    }

    function browserCondition() {
        return appIf(['^com\\.google\\.Chrome.*$', '^com\\.microsoft\\.edgemac.*$']);
    }

    function chromeCondition() {
        return appIf(['^com\\.google\\.Chrome.*$']);
    }

    function excelCondition() {
        return appIf(['^com\\.microsoft\\.Excel$']);
    }

    function excelUnlessCondition() {
        return appUnless(['^com\\.microsoft\\.Excel$']);
    }

    function finderCondition() {
        return appIf(['^com\\.apple\\.finder$']);
    }

    function macTextAppCondition() {
        return appIf(['^com\\.apple\\.TextEdit$', '^com\\.apple\\.Notes$']);
    }

    function jeffModeCondition() {
        return { type: 'variable_if', name: 'jeff_sheets_mode', value: true };
    }

    function jeffModeOffCondition() {
        return { type: 'variable_unless', name: 'jeff_sheets_mode', value: true };
    }

    function chromeSheetOrTabCommand(previous) {
        var sheetsKeyCode = previous ? 126 : 125;
        return [
            '/usr/bin/osascript',
            '-e \'tell application "Google Chrome" to set currentURL to URL of active tab of front window\'',
            '-e \'tell application "System Events"\'',
            '-e \'if currentURL starts with "https://docs.google.com/spreadsheets/" then\'',
            '-e \'key code ' + sheetsKeyCode + ' using {option down}\'',
            '-e \'else\'',
            previous
                ? '-e \'key code 48 using {control down, shift down}\''
                : '-e \'key code 48 using {control down}\'',
            '-e \'end if\'',
            '-e \'end tell\''
        ].join(' ');
    }

// ============================================================
// Alt+Down
//
// Google Sheets:
//   Ctrl+Command+R -> フィルターメニューを開く
//
// その他Chrome:
//   通常のOption+Down
// ============================================================
function chromeFilterDropdownOrAltDownCommand() {

    return [
        '/usr/bin/osascript',

        '-e \'tell application "Google Chrome" to set currentURL to URL of active tab of front window\'',

        '-e \'tell application "System Events"\'',

        '-e \'if currentURL starts with "https://docs.google.com/spreadsheets/" then\'',

        '-e \'keystroke "r" using {control down, command down}\'',

        '-e \'else\'',

        '-e \'key code 125 using {option down}\'',

        '-e \'end if\'',

        '-e \'end tell\''
    ].join(' ');
}

    function keyRule(fromKey, mandatory, toKey, toModifiers, conditions, optional) {
        return {
            type: 'basic',
            from: {
                key_code: fromKey,
                modifiers: {
                    mandatory: mandatory || [],
                    optional: optional || ['caps_lock']
                }
            },
            to: [{ key_code: toKey, modifiers: toModifiers || [] }],
            conditions: conditions
        };
    }

    var manipulators = [];

    // Excel QAT: Alt+1/2/3 -> PERSONAL.XLSB bridge shortcuts.
    manipulators.unshift(keyRule('1', ['option'], 'y', ['left_control','left_shift'], [deviceCondition(), excelCondition()]));
    manipulators.unshift(keyRule('2', ['option'], 'g', ['left_control','left_shift'], [deviceCondition(), excelCondition()]));
    manipulators.unshift(keyRule('3', ['option'], 'x', ['left_control','left_shift'], [deviceCondition(), excelCondition()]));

// ============================================================
// Excel for Mac - Jeff Alt shortcuts
// ============================================================

// Ctrl + Alt + W
// → Jeff: DefaultSquareWidth
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

// Ctrl + Shift + Alt + F
// → Jeff: SetToDefaultFont
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

    // Excel Ctrl+Arrow / Ctrl+Shift+Arrow.
    ['up_arrow','down_arrow','left_arrow','right_arrow'].forEach(function (key) {
        manipulators.push(keyRule(key, ['control'], key, ['left_command'], [deviceCondition(), excelCondition()]));
        manipulators.push(keyRule(key, ['control','shift'], key, ['left_command','left_shift'], [deviceCondition(), excelCondition()]));
    });

    // TextEdit / Notes Ctrl+Shift+V.
    manipulators.unshift(keyRule('v', ['control','shift'], 'v', ['left_command','left_option','left_shift'], [deviceCondition(), macTextAppCondition()]));


// ============================================================
// Altキー
// 1) Alt+Tab -> Command+Tab
// 2) Altを単独で押して離したら Windows KeyTips 開始
// ============================================================
['left_option', 'right_option'].forEach(function (optionKey) {
    var commandKey =
        optionKey === 'left_option' ? 'left_command' : 'right_command';

    manipulators.push({
        type: 'basic',

        from: {
            key_code: optionKey,
            modifiers: {
                optional: ['any']
            }
        },

        // Altを他のキーと組み合わせた場合はOptionとして普通に使う
        to: [
            {
                key_code: optionKey
            }
        ],

        // Altだけを押して離した場合:
        // 2秒間、WindowsのAltメニュー入力待ち状態にする
        to_if_alone: [
            {
                set_variable: {
                    name: 'windows_alt_expiration',
                    expression: 'system.now.milliseconds + 2000'
                }
            }
        ],

        // Alt+TabだけはCommand+Tabへ
        to_if_other_key_pressed: [
            {
                other_keys: [
                    {
                        key_code: 'tab',
                        modifiers: {
                            optional: ['any']
                        }
                    }
                ],
                to: [
                    {
                        key_code: commandKey
                    }
                ]
            }
        ],

        conditions: [
            deviceCondition()
        ],

        parameters: {
            'basic.to_if_alone_timeout_milliseconds': 500
        }
    });
});

    // General Windows Ctrl shortcuts -> Command. Excel is excluded.
    var ctrlToCommandKeys = [
        'a','b','c','d','f','g','i','k','l','n','o','p','r','s','t','u','v','w','x','z',
        '0','1','2','3','4','5','6','7','8','9','equal_sign','hyphen'
    ];
    ctrlToCommandKeys.forEach(function (key) {
        manipulators.push(keyRule(
            key, ['control'], key, ['left_command'],
            [deviceCondition(), excelUnlessCondition()],
            ['shift','caps_lock']
        ));
    });

// ============================================================
// Windows Excel風 Alt -> A -> T
// ============================================================
//
// 操作:
// Altを押して離す
// → A
// → T
//
// Excel:
//   Ctrl+Shift+L に変換してフィルターON/OFF
//
// Google Sheets:
//   Ctrl+Option+Shift+L をChrome拡張へ送る
//   （Chrome拡張側にFILTER_TOGGLEを追加する）
// ============================================================


// ------------------------------------------------------------
// Step 1: Alt の次に A
// ------------------------------------------------------------
manipulators.unshift({
    type: 'basic',

    from: {
        key_code: 'a',
        modifiers: {
            optional: ['caps_lock']
        }
    },

    to: [
        {
            set_variable: {
                name: 'windows_alt_a_expiration',
                expression: 'system.now.milliseconds + 2000'
            }
        },
        {
            set_variable: {
                name: 'windows_alt_expiration',
                value: 0
            }
        }
    ],

    conditions: [
        deviceCondition(),
        excelCondition(),
    
        {
            type: 'expression_if',
            expression:
                'windows_alt_expiration > system.now.milliseconds'
        }
    ]
});


// ------------------------------------------------------------
// Step 2: Excelで T
// Alt -> A -> T
// → Ctrl+Shift+L
// → フィルター追加/解除
// ------------------------------------------------------------
manipulators.unshift({
    type: 'basic',

    from: {
        key_code: 't',
        modifiers: {
            optional: ['caps_lock']
        }
    },

    to: [
        {
            key_code: 'l',
            modifiers: [
                'left_control',
                'left_shift'
            ]
        },
        {
            set_variable: {
                name: 'windows_alt_a_expiration',
                value: 0
            }
        }
    ],

    conditions: [
        deviceCondition(),
        excelCondition(),

        {
            type: 'expression_if',
            expression:
                'windows_alt_a_expiration > system.now.milliseconds'
        }
    ]
});

    // Chrome / Edge Windows-like corrections.
    manipulators.unshift(keyRule('h', ['control'], 'y', ['left_command'], [deviceCondition(), browserCondition()]));
    manipulators.unshift(keyRule('j', ['control'], 'j', ['left_command','left_shift'], [deviceCondition(), browserCondition()]));
    manipulators.unshift(keyRule('u', ['control'], 'u', ['left_command','left_option'], [deviceCondition(), browserCondition()]));
    manipulators.unshift(keyRule('i', ['control','shift'], 'i', ['left_command','left_option'], [deviceCondition(), browserCondition()]));
    manipulators.unshift(keyRule('c', ['control','shift'], 'c', ['left_command','left_option'], [deviceCondition(), browserCondition()]));

    // Alt+F4 / Ctrl+F4 / Ctrl+Shift+Esc.
    manipulators.push(keyRule('f4', ['option'], 'w', ['left_command'], [deviceCondition()]));
    manipulators.push(keyRule('f4', ['control'], 'w', ['left_command'], [deviceCondition()]));
    manipulators.push(keyRule('escape', ['control','shift'], 'escape', ['left_command','left_option'], [deviceCondition()]));

    // Japanese input.
    manipulators.push({
        type: 'basic', from: { key_code: 'grave_accent_and_tilde' },
        to: [{ key_code: 'japanese_eisuu' }],
        conditions: [deviceCondition(), { type: 'input_source_if', input_sources: [{ language: '^ja$' }] }]
    });
    manipulators.push({
        type: 'basic', from: { key_code: 'grave_accent_and_tilde' },
        to: [{ key_code: 'japanese_kana' }],
        conditions: [deviceCondition(), { type: 'input_source_unless', input_sources: [{ language: '^ja$' }] }]
    });
    manipulators.push({
        type: 'basic',
        from: { key_code: 'japanese_pc_katakana', modifiers: { optional: ['any'] } },
        to: [{ key_code: 'japanese_kana' }],
        conditions: [deviceCondition()]
    });

    // Finder Alt+Up.
    manipulators.unshift(keyRule('up_arrow', ['option'], 'up_arrow', ['left_command'], [deviceCondition(), finderCondition()]));

    // Finder Ctrl+X.
    manipulators.unshift({
        type: 'basic',
        from: { key_code: 'x', modifiers: { mandatory: ['control'], optional: ['caps_lock'] } },
        to: [
            { key_code: 'c', modifiers: ['left_command'] },
            { set_variable: { name: 'finder_cut_pending', value: true } }
        ],
        conditions: [deviceCondition(), finderCondition()]
    });

    // Finder Ctrl+C.
    manipulators.unshift({
        type: 'basic',
        from: { key_code: 'c', modifiers: { mandatory: ['control'], optional: ['caps_lock'] } },
        to: [
            { key_code: 'c', modifiers: ['left_command'] },
            { set_variable: { name: 'finder_cut_pending', value: false } }
        ],
        conditions: [deviceCondition(), finderCondition()]
    });

    // Finder Ctrl+V after Ctrl+X.
    manipulators.unshift({
        type: 'basic',
        from: { key_code: 'v', modifiers: { mandatory: ['control'], optional: ['caps_lock'] } },
        to: [
            { key_code: 'v', modifiers: ['left_command','left_option'] },
            { set_variable: { name: 'finder_cut_pending', value: false } }
        ],
        conditions: [
            deviceCondition(), finderCondition(),
            { type: 'variable_if', name: 'finder_cut_pending', value: true }
        ]
    });

    // Finder normal Ctrl+V.
    manipulators.unshift({
        type: 'basic',
        from: { key_code: 'v', modifiers: { mandatory: ['control'], optional: ['caps_lock'] } },
        to: [{ key_code: 'v', modifiers: ['left_command'] }],
        conditions: [
            deviceCondition(), finderCondition(),
            { type: 'variable_unless', name: 'finder_cut_pending', value: true }
        ]
    });

    // Finder Enter -> open except text editing.
    manipulators.unshift({
        type: 'basic',
        from: { key_code: 'return_or_enter', modifiers: { optional: ['caps_lock'] } },
        to: [{ key_code: 'o', modifiers: ['left_command'] }],
        conditions: [
            deviceCondition(), finderCondition(),
            { type: 'expression_unless', expression: "accessibility.focused_ui_element.role_string like 'AXText*'" }
        ]
    });

    // Finder F2 / Delete / Shift+Delete.
    manipulators.push({
        type: 'basic', from: { key_code: 'f2' }, to: [{ key_code: 'return_or_enter' }],
        conditions: [deviceCondition(), finderCondition()]
    });
    manipulators.push({
        type: 'basic', from: { key_code: 'delete_forward' },
        to: [{ key_code: 'delete_or_backspace', modifiers: ['left_command'] }],
        conditions: [deviceCondition(), finderCondition()]
    });
    manipulators.push(keyRule('delete_forward', ['shift'], 'delete_or_backspace', ['left_command','left_option'], [deviceCondition(), finderCondition()]));

    // Windows key shortcuts.
    manipulators.push(keyRule('l', ['command'], 'q', ['left_control','left_command'], [deviceCondition()]));
    manipulators.push({
        type: 'basic',
        from: { key_code: 'e', modifiers: { mandatory: ['command'], optional: ['caps_lock'] } },
        to: [{ software_function: { open_application: { bundle_identifier: 'com.apple.finder' } } }],
        conditions: [deviceCondition()]
    });
    manipulators.push(keyRule('r', ['command'], 'spacebar', ['left_command'], [deviceCondition()]));
    manipulators.push(keyRule('s', ['command','shift'], '4', ['left_control','left_shift','left_command'], [deviceCondition()]));
    manipulators.push({
        type: 'basic', from: { key_code: 'print_screen' },
        to: [{ key_code: '3', modifiers: ['left_control','left_shift','left_command'] }],
        conditions: [deviceCondition()]
    });

    // Browser F5 / F12.
    manipulators.push(keyRule('f5', [], 'r', ['left_command'], [deviceCondition(), browserCondition()]));
    manipulators.push(keyRule('f12', [], 'i', ['left_command','left_option'], [deviceCondition(), browserCondition()]));

    // =========================================================================
    // Jeff Google Sheets mode
    // Toggle: Ctrl+Option+Shift+0
    // =========================================================================
    var jeffManipulators = [];

// ============================================================
// Google Sheets: Alt+Down
// → フィルターのドロップダウンを開く
// ============================================================
jeffManipulators.unshift({
    type: 'basic',

    from: {
        key_code: 'down_arrow',
        modifiers: {
            mandatory: ['option'],
            optional: ['caps_lock']
        }
    },

    to: [
        {
            shell_command:
                chromeFilterDropdownOrAltDownCommand()
        }
    ],

    conditions: [
        deviceCondition(),
        chromeCondition(),
        jeffModeCondition()
    ]
});

    jeffManipulators.push({
        type: 'basic',
        from: { key_code: '0', modifiers: { mandatory: ['control','option','shift'], optional: ['caps_lock'] } },
        to: [
            { set_variable: { name: 'jeff_sheets_mode', value: true } },
            { set_notification_message: { id: 'jeff_sheets_mode', text: "Jeff's Sheets mode: ON" } }
        ],
        conditions: [deviceCondition(), jeffModeOffCondition()]
    });

    jeffManipulators.push({
        type: 'basic',
        from: { key_code: '0', modifiers: { mandatory: ['control','option','shift'], optional: ['caps_lock'] } },
        to: [
            { set_variable: { name: 'jeff_sheets_mode', value: false } },
            { set_notification_message: { id: 'jeff_sheets_mode', text: "Jeff's Sheets mode: OFF" } }
        ],
        conditions: [deviceCondition(), jeffModeCondition()]
    });

    // Ctrl+PageUp/PageDown:
    // Sheets -> sheet move, other Chrome pages -> tab move.
    jeffManipulators.unshift({
        type: 'basic',
        from: { key_code: 'page_up', modifiers: { mandatory: ['control'], optional: ['caps_lock'] } },
        to: [{ shell_command: chromeSheetOrTabCommand(true) }],
        conditions: [deviceCondition(), chromeCondition(), jeffModeCondition()]
    });
    jeffManipulators.unshift({
        type: 'basic',
        from: { key_code: 'page_down', modifiers: { mandatory: ['control'], optional: ['caps_lock'] } },
        to: [{ shell_command: chromeSheetOrTabCommand(false) }],
        conditions: [deviceCondition(), chromeCondition(), jeffModeCondition()]
    });

    // Google Sheets Alt+1 is NOT intercepted here.
    // The Chrome content script should directly map A+Digit1 -> PRECEDENT.

    function addJeffPassthrough(key, mandatory) {
        var out = [];
        mandatory.forEach(function (m) {
            if (m === 'control') out.push('left_control');
            if (m === 'option') out.push('left_option');
            if (m === 'shift') out.push('left_shift');
            if (m === 'command') out.push('left_command');
        });
        jeffManipulators.push(keyRule(
            key, mandatory, key, out,
            [deviceCondition(), chromeCondition(), jeffModeCondition()]
        ));
    }

    // Jeff formatting / colors.
    addJeffPassthrough('n', ['control','shift']);
    addJeffPassthrough('c', ['control','shift']);
    addJeffPassthrough('1', ['control','shift']);
    addJeffPassthrough('5', ['control','shift']);
    addJeffPassthrough('c', ['control','option','shift']);
    addJeffPassthrough('v', ['control','option','shift']);
    addJeffPassthrough('w', ['control','shift']);
    addJeffPassthrough('m', ['control','shift']);
    addJeffPassthrough('m', ['control','option','shift']);
    addJeffPassthrough('t', ['control','shift']);
    addJeffPassthrough('i', ['control','shift']);
    addJeffPassthrough('f', ['control','option','shift']);
    addJeffPassthrough('b', ['control','shift']);
    addJeffPassthrough('b', ['control','option','shift']);

    // Jeff SuperFill / formulas.
    addJeffPassthrough('d', ['control','shift']);
    addJeffPassthrough('d', ['control','option']);
    addJeffPassthrough('z', ['control','shift']);
    addJeffPassthrough('r', ['control','shift']);
    addJeffPassthrough('r', ['control','option']);
    addJeffPassthrough('q', ['control','shift']);
    addJeffPassthrough('f4', ['control','shift']);
    addJeffPassthrough('f4', ['control','option']);
    addJeffPassthrough('e', ['control','shift']);
    addJeffPassthrough('e', ['control','option','shift']);

    // Sheets Ctrl+Arrow -> Command+Arrow.
    ['up_arrow','down_arrow','left_arrow','right_arrow'].forEach(function (key) {
        jeffManipulators.push(keyRule(
            key, ['control'], key, ['left_command'],
            [deviceCondition(), chromeCondition(), jeffModeCondition()]
        ));
    });

    // Jeff dimensions / alignment / decimals / save / zoom.
    addJeffPassthrough('w', ['control','option','shift']);
    addJeffPassthrough('h', ['control','option','shift']);
    addJeffPassthrough('right_arrow', ['control','option','shift']);
    addJeffPassthrough('left_arrow', ['control','option','shift']);
    addJeffPassthrough('down_arrow', ['control','option','shift']);
    addJeffPassthrough('up_arrow', ['control','option','shift']);
    addJeffPassthrough('w', ['control','option']);
    addJeffPassthrough('l', ['control','option']);
    addJeffPassthrough('h', ['control','option']);
    addJeffPassthrough('comma', ['control','shift']);
    addJeffPassthrough('period', ['control','shift']);
    addJeffPassthrough('s', ['control','shift']);
    addJeffPassthrough('o', ['control','option','shift']);
    addJeffPassthrough('k', ['control','shift']);
    addJeffPassthrough('j', ['control','shift']);
    addJeffPassthrough('h', ['control','shift']);
    addJeffPassthrough('1', ['control','option']);
    addJeffPassthrough('f4', ['control','option','shift']);

    manipulators = jeffManipulators.concat(manipulators);

    return {
        description: 'Y-U0036 - Windows-like shortcuts',
        manipulators: manipulators
    };
}

main();