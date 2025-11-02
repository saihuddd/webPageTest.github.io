// 班次替换规则
const ReplaceRules = {
    rules: {
        "备1": "备1-可休",
        "十四": "十四 8.00~16.45",
        "夜": "夜 21.30~7.30",
        "晚": "晚 15.00~22.30",
        "早": "早 7.30~15.00",
        "中": "中 15.00~21.30",
        "四": "四 8.00~16.45",
        "五": "五 8.00~12.00",
        "二": "二 8.00~16.45",
        "三": "三 8.00~16.45",
        "九": "九 8.00~16.45",
        "十": "十 8.00~16.45",
        "1": "1 8.00~16.45",
        "2": "2 7.00-15.00",
        "3": "3 7.00-15.00",
        "4": "4 7.30~17.15",
        "5": "5 8.00~17.15",
        "6": "6 7.30~16.45",
        "7": "7 8.30~17.15",
        "9": "9 7.30~16.55",
        "15": "15 8.00~12.00",
        "西": "西 8.00~12.00",
        "备": "备 7.30~16.30",
        "帮": "帮 8.30~16.15",
        "休": "休息",
        "工": "工休"
    },

    // 应用替换规则
    applyRules(text) {
        if (text === undefined || text === null) {
            return '';
        }

        let result = String(text);
        const escapeRegExp = value => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        Object.entries(this.rules)
            .sort((a, b) => b[0].length - a[0].length) // 长的规则先替换
            .forEach(([key, value]) => {
                const pattern = `(?<!\\S)${escapeRegExp(key)}(?!\\S)`;
                const regex = new RegExp(pattern, 'g');
                result = result.replace(regex, value);
            });
        return result;
    },

    // 保存规则到本地存储
    saveRules() {
        localStorage.setItem('replaceRules', JSON.stringify(this.rules));
    },

    // 从本地存储加载规则
    loadRules() {
        const saved = localStorage.getItem('replaceRules');
        if (saved) {
            this.rules = JSON.parse(saved);
        }
    }
};

window.ReplaceRules = ReplaceRules;