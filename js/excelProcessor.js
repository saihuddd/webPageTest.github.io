// Excel 处理模块
const ExcelProcessor = {
    parseDate(rawDate, year, month) {
        const value = (rawDate ?? '').toString().trim();
        if (!value) {
            return '';
        }

        console.log('[parseDate] 开始解析日期:', { rawDate, value, year, month });

        // 优先处理：如果只有1-2位数字，直接使用传入的年份和月份
        if (/^\d{1,2}$/.test(value) && year && month) {
            const normalizedMonth = String(month).padStart(2, '0');
            const normalizedDay = value.padStart(2, '0');
            const result = `${year}-${normalizedMonth}-${normalizedDay}`;
            console.log('[parseDate] 匹配为纯数字日期:', result);
            return result;
        }

        const normalizedValue = value
            .replace(/[\u3000\s]+/g, '')
            .replace(/年|\.|\//g, '-');

        // 处理包含年份的格式
        const candidateFormats = [
            'YYYY-MM-DD',
            'YYYY-M-D',
            'MM-DD-YYYY',
            'M-D-YYYY',
            'YYYY年M月D日',
            'YYYY年MM月DD日'
        ];

        for (const format of candidateFormats) {
            const parsed = dayjs(value, format, true);
            if (parsed.isValid()) {
                const result = parsed.format('YYYY-MM-DD');
                console.log('[parseDate] 匹配格式:', format, '->', result);
                return result;
            }
        }

        // 处理只有月日的格式，强制使用传入的年份
        const monthDayFormats = ['M-D', 'MM-DD', 'M月D日', 'MM月DD日'];
        for (const format of monthDayFormats) {
            const parsed = dayjs(value, format, true);
            if (parsed.isValid() && year) {
                const parsedMonth = parsed.month() + 1; // dayjs month is 0-based
                const parsedDay = parsed.date();
                const result = `${year}-${String(parsedMonth).padStart(2, '0')}-${String(parsedDay).padStart(2, '0')}`;
                console.log('[parseDate] 匹配月日格式:', format, '->', result);
                return result;
            }
        }

        // 处理标准化后的值
        const normalizedParsed = dayjs(normalizedValue, ['YYYY-M-D', 'YYYYMMDD'], true);
        if (normalizedParsed.isValid()) {
            const result = normalizedParsed.format('YYYY-MM-DD');
            console.log('[parseDate] 匹配标准化格式:', result);
            return result;
        }

        // 处理 Excel 序列号（日期序列号通常 > 1，但可能小于 20000）
        if (!Number.isNaN(Number(value))) {
            const numericDate = Number(value);
            // Excel 日期序列号从 1900-01-01 开始，1 = 1900-01-01
            // 所以即使是小数字也可能是日期序列号
            if (numericDate >= 1 && numericDate < 1000000) {
                // 尝试使用 XLSX 的日期解析
                if (XLSX?.SSF?.parse_date_code) {
                    try {
                        const parsedExcelDate = XLSX.SSF.parse_date_code(numericDate);
                        if (parsedExcelDate) {
                            const { y, m, d } = parsedExcelDate;
                            const excelDate = dayjs(`${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`, 'YYYY-MM-DD', true);
                            if (excelDate.isValid()) {
                                const result = excelDate.format('YYYY-MM-DD');
                                console.log('[parseDate] 匹配Excel序列号:', numericDate, '->', result);
                                return result;
                            }
                        }
                    } catch (e) {
                        console.warn('[parseDate] Excel序列号解析失败:', e);
                    }
                }
            }
        }

        // 如果传入年份和月份，尝试从字符串中提取数字并组合
        if (year && month) {
            const sanitized = value.replace(/[^0-9]/g, '');
            if (sanitized.length <= 2 && sanitized.length > 0) {
                const normalizedMonth = String(month).padStart(2, '0');
                const normalizedDay = sanitized.padStart(2, '0');
                const result = `${year}-${normalizedMonth}-${normalizedDay}`;
                console.log('[parseDate] 使用传入年月组合:', result);
                return result;
            }
        }

        // 最后尝试：dayjs 自动解析（但需要验证年份是否合理）
        const directParsed = dayjs(value);
        if (directParsed.isValid()) {
            const parsedYear = directParsed.year();
            // 如果解析出的年份不合理（比如2001年），且我们传入了年份，则使用传入的年份
            if (parsedYear < 2000 || parsedYear > 2100) {
                if (year) {
                    const result = `${year}-${String(directParsed.month() + 1).padStart(2, '0')}-${String(directParsed.date()).padStart(2, '0')}`;
                    console.log('[parseDate] 修正不合理年份:', directParsed.format('YYYY-MM-DD'), '->', result);
                    return result;
                }
            } else {
                const result = directParsed.format('YYYY-MM-DD');
                console.log('[parseDate] dayjs自动解析:', result);
                return result;
            }
        }

        console.warn('[parseDate] 无法解析日期:', { rawDate, value, year, month });
        return '';
    },

    async processFile(file, { targetName, year, month }) {
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            const rows = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: false
            });

            if (rows.length < 2) {
                throw new Error('Excel 文件缺少表头');
            }

            const headerRow = rows[1].map(cell => (cell ?? '').toString().trim());
            const sanitize = value => value.replace(/\s+/g, '');

            const dateIndex = headerRow.findIndex(col => sanitize(col) === '日期');
            if (dateIndex === -1) {
                throw new Error('未找到 "日期" 列');
            }

            const weekdayIndex = headerRow.findIndex(col => sanitize(col) === '星期');

            const targetIndex = headerRow.findIndex(col => sanitize(col) === sanitize(targetName));
            if (targetIndex === -1) {
                throw new Error(`未在 Excel 中找到姓名列: ${targetName}`);
            }

            const scheduleData = {};
            const remarks = [];

            for (let i = 2; i < rows.length; i += 1) {
                const row = rows[i];
                if (!row) {
                    continue;
                }

                const rawDate = row[dateIndex];
                const weekday = weekdayIndex !== -1 ? (row[weekdayIndex] ?? '').toString().trim() : '';
                let content = (row[targetIndex] ?? '').toString().trim();

                if (!rawDate && !content) {
                    continue;
                }

                if (!content) {
                    remarks.push(ExcelProcessor.composeRemark(row));
                    continue;
                }

                content = ReplaceRules.applyRules(content);

                const parsedDate = ExcelProcessor.parseDate(rawDate, year, month);

                if (!parsedDate) {
                    remarks.push(ExcelProcessor.composeRemark(row));
                    continue;
                }

                const dateObj = dayjs(parsedDate);
                if (!dateObj.isValid()) {
                    remarks.push(ExcelProcessor.composeRemark(row));
                    continue;
                }

                const dateStr = dateObj.format('YYYY-MM-DD');
                const entry = scheduleData[dateStr];
                const lunar = LunarCalendar.getLunarDate(dateObj.toDate());
                const holiday = LunarCalendar.getHoliday(dateStr);

                if (entry) {
                    entry.content = `${entry.content} | ${content}`;
                } else {
                    scheduleData[dateStr] = {
                        weekday,
                        content,
                        lunar,
                        holiday
                    };
                }
            }

            // 打印填充后的字典内容（用于调试）
            console.log('=== 排班字典内容 ===');
            console.log('文件:', file.name);
            console.log('字典条目数:', Object.keys(scheduleData).length);
            console.log('字典内容:', scheduleData);
            
            // 按日期排序打印，方便查看
            const sortedEntries = Object.entries(scheduleData).sort((a, b) => a[0].localeCompare(b[0]));
            console.log('按日期排序的条目:');
            sortedEntries.forEach(([dateStr, info]) => {
                console.log(`  ${dateStr}: ${info.content || '(空)'}`);
            });
            
            if (remarks.length > 0) {
                console.log('备注条目数:', remarks.length);
                console.log('备注内容:', remarks);
            }
            console.log('==================');

            return { scheduleData, remarks };
        } catch (error) {
            console.error('Excel处理错误:', error);
            throw error;
        }
    },

    composeRemark(row) {
        return row
            .filter(cell => cell !== undefined && cell !== null && String(cell).trim() !== '')
            .map(cell => String(cell).trim())
            .join('\t');
    }
};

window.ExcelProcessor = ExcelProcessor;