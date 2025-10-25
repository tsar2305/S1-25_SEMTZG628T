exports.EmailPage =
    class EmailPage {
        constructor(page) {
            this.page = page;
            this.emailUrl = "https://www.mailinator.com";
            this.username = "//input[@id='search']";
            this.inboxTable = "table[class='table-striped jambo_table'] tbody";
            this.textFrame="[id='html_msg_body']";
            this.codes="(//td[contains(@class,'bui-mobile')]//strong)";
            this.hotelHeader = "(//h1)[2]";
            this.hotelAddress = "(//td[contains(@class,'bui-mobile')])//address";
        }

        async rowFilterSelect(rows, value) {
            const matchRow = rows.filter({
                has: await this.page.locator('td'),
                hasText: value
            });
            await matchRow.click();
            return matchRow;
        };
        async navigateToPage(url) {
            await this.page.goto(url);
            await this.page.waitForLoadState('load');
        };
        async selectText(value) {
            await (await this.page.waitForSelector("//*[text()='" + value + "']")).click();
        }
        async getInbox() {
            var arr = [];
            await this.page.waitForSelector(this.inboxTable);
            const table = await this.page.locator(this.inboxTable);
            var rows = await table.locator("tr");
            await this.rowFilterSelect(rows, 'Your booking is confirmed');
            await this.page.waitForLoadState('load');
            await this.page.waitForSelector(this.textFrame);
            const frame = await this.page.frameLocator(this.textFrame);
            arr.push(await frame.locator(this.codes+'[1]').textContent());
            arr.push(await frame.locator(this.codes+'[2]').textContent());
            arr.push(await frame.locator(this.hotelHeader).textContent());
            arr.push(await frame.locator(this.hotelAddress).textContent());
            console.log(arr);
            return arr;
        }
    }