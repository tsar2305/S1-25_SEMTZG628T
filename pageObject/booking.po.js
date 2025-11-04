const { randomInt } = require("node:crypto");
const { faker, fakerEN_IN, fakerEN_US, simpleFaker, fakerBASE, fakerEN_CA } = require('@faker-js/faker');
const { error } = require("node:console");
const { errors } = require("playwright-extra");
const { throws } = require("node:assert");
var attempt=0;
exports.BookingPage =
    class BookingPage {

        constructor(page) {
            this.page = page;
            this.bookingUrl = "https://www.booking.com/";
            this.userName = "//input[@name='username']";
            this.submitButton = "//button[@type='submit']";
            this.otpInput = "#otp_input";
            this.logo = "div[class='Header_logo']";
            this.currencyBtn = "(//button[@type='button'])[1]";
            this.currencyValues = "//span[@class='Picker_selection-text']";
            this.destination = "//*[@data-testid ='destination-container']//input";
            this.autoLocations = "//div[@id='autocomplete-results']//ul/li";
            this.datePicker = "//button[@data-testid='searchbox-dates-container']";
            this.nextDateButton = "//button[@aria-label='Next month']";
            this.currencyFrame = "(//iframe)[7]";
            this.currentDate = "(//h3[@aria-live='polite'])[1]";
            this.dates = "(//table[@role='grid']//tr/td/span)";
            this.signInClose = "button[aria-label='Dismiss sign-in info.']";
            this.occupancyButton = "button[data-testid='occupancy-config']";
            this.cntrlButton = "(//button[@tabindex='-1'])";
            this.occupancyValue = "(//span[@aria-hidden='true'])";
            this.searchResultHeader = "//*[contains(@aria-label,'Search results updated')]";
            this.sortButton = "//*[contains(@data-testid,'dropdown-trigger')]";
            this.Nopaymnt = "(//div[@data-testid='filters-group-container']//*[contains(text(),'No prepayment')])[1]";
            this.NoCredCrd = "(//div[@data-testid='filters-group-container']//*[contains(text(),'Book without credit card')])[1]";
            this.showDrpdwn = "(//div[contains(text(),'Show ')])";
            this.usedFilters1 = "(//button[contains(@data-testid,'filter:')])";
            this.usedFilters2 = "((//div[@data-testid='filters-group-container'])[1]//div[contains(@data-filters-item,'used_filters:')]//div[@data-testid='filters-group-label-content'])";
            this.availablityBtn = "(//*[@data-testid = 'availability-cta-btn'])[1]";
            this.availableBtn = "//*[@data-testid = 'availability-cta-btn']";
            this.reserve = "//button[contains(@id,'book')]";
            this.checkboxes = "(//*[@type='checkbox'])";
            this.filterGrouped = "//button[contains(@data-testid,'filter-group:')]";
            this.hotelHeader = "(//h2)[1]";
            this.ruleContent = "//div[text()='House rules']/following::div[@data-testid='property-section--content']//*[text()]";
            this.firstName = "(//input[@name='firstname'])[1]";
            this.lastName = "(//input[@name='lastname'])[1]";
            this.email = "(//input[@name='email'])[1]";
            this.confirmEmail = "(//input[@name='email_confirmation'])[1]";
            this.country = "//select[@name='cc1']";
            this.city = "(//input[@name='city'])[1]";
            this.zip = "(//input[@name='zip'])[1]";
            this.phone = "(//input[@name='phoneNumber'])[1]";
            this.finalBtn = "(//button[@name='book'])[1]";
            this.bookadres = "(//button[@type='button'])[3]//div";
            this.recommendSubmit = "//*[contains(@class,'recommendation-reserve-btn')]";
            this.cnfrmSubmit = "//*[@class='bui-button__text js-reservation-button__text']";
            this.recommendContent = "(//*[@class='e2e-gr-room-name room_link'])[1]";
            this.recommendFor = "[class='bui-price-display__label ']";
            this.recommendPrice = "(//*[@class='prco-valign-middle-helper'])[1]";
            this.pricetax = "(//div[contains(@class,'prd-taxes-and-fees-under-price')])[1]";
            this.hotelTable = "[id='hprt-table']";
            this.maxOccupancy = "//*[@id='hprt-table']//tr/td[1]//span[2]";
            this.tableRoom = "(//*[@id='hprt-table']//tr/th//span[@class='hprt-roomtype-icon-link '])[1]";
            this.agreeBox = "//*[@name='booking_bv']";
            this.popUpCls = "button[aria-label='Close modal']";
            this.confirmation = "(//div[@class='reservation-status']//..//span)";
            this.header = "//div[@class='confirmation_status_container']//..//h1";
            this.checkboxContent = "//*[@data-testid='filters-group-label-content']";
            this.review = "//div[@data-testid='review-score']//div[2]";
            this.hotels = "//div//div[@data-testid='property-card-container']";
            this.pricing = "//span[@data-testid='price-and-discounted-price']";
            this.paymnt = "(//span[contains(text(),'Payment terms')])[2]";
            this.address = "(//input[@name='address1'])[1]";
            this.city = "(//input[@name='city'])[1]";
            this.loadMore = "//*[contains(text(),'Load more result')]";
        }
        async monthNo(month) {
            var month_short = month;
            const month_number = {
                "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
                "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
                "Sept": 9, "Oct": 10, "Nov": 11, "Dec": 12
            }[month_short];
            return month_number;
        }
        async ranData() {
            const data = [];
            const amenities = ['Wifi', 'Parking', 'Restraunt', 'Swimming Pool', 'Spa', 'Room Service', 'shuttle', 'Wheelchair', '24-hour front desk'];
            const currency = ['INR', 'USD', 'ARS', 'AUD', 'CNY', 'KWD', 'QAR', 'TWD', 'NZD', 'THB'];
            const count = faker.number.int({ min: 1, max: 3 });
            const picked = faker.helpers.arrayElements(amenities, count);
            const filterValue = picked.join(',');
            data.push(faker.person.firstName());
            data.push(faker.person.lastName());
            data.push(data[0] + faker.number.int({ min: 11, max: 99 }) + '@mailinator.com');
            data.push(faker.phone.number({ style: 'national' }));
            data.push(faker.location.country());
            data.push(filterValue);
            data.push(currency[faker.number.int({ min: 0, max: 9 })]);
            data.push(faker.number.int({ min: 1, max: 5 })); //member
            data.push(faker.number.int({ min: parseInt(data[7] / 2) < 2 ? 1 : (data[7] / 2), max: data[7] - 1 <= 2 ? 2 : data[7] - 1 }));//room
            const currentDate = new Date();
            var arrivalDay = faker.number.int({ min: currentDate.getDate(), max: 28 });
            var departureDay = arrivalDay + faker.number.int({ min: 1, max: 6 });
            data.push(Math.floor(randomInt(5000, 15000) / 1000) * 1000 * data[8] * (departureDay - arrivalDay));//budget
            var month = await this.monthShrt(currentDate.getMonth() + 1);
            var year = await currentDate.getFullYear();
            var arrival = arrivalDay + '.' + month + '.' + year;
            if (currentDate.getMonth() == 12 && arrivalDay + 5 > 31) {
                year = await currentDate.getFullYear() + 1;
            }
            if ((currentDate.getMonth() + 1) == 12 && arrivalDay + 5 > 31) {
                month = await this.monthShrt(1);
            } else if (arrivalDay + 2 > 31) {
                month = await this.monthShrt(currentDate.getMonth() + 2);
            }
            var departure = departureDay + '.' + month + '.' + year;
            data.push(arrival);
            data.push(departure);
            data.push(fakerEN_IN.location.city());
            console.log(data);
            return data;
        }
        async monthShrt(no) {
            const month_short = {
                1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr",
                5: "May", 6: "Jun", 7: "Jul", 8: "Aug",
                9: "Sept", 10: "Oct", 11: "Nov", 12: "Dec"
            }[no];
            return month_short;
        }
        async navigateToPage(url) {
            await this.page.goto(url);
            await this.page.reload();
            await this.page.waitForLoadState('load');
        }
        async closeSignIn() {
            let close = this.page.locator(this.signInClose);
            if (await close.isVisible()) {
                await close.click();
            }
        };
        async selectDestination(location) {
            await this.page.waitForLoadState('load');
            await this.page.waitForSelector(this.destination);
            await this.closeSignIn();
            try {
                await this.page.locator(this.destination).click();
                await this.page.locator(this.destination).clear();
                await this.page.locator(this.destination).fill(location);
                await this.page.waitForTimeout(3000);
                await this.autocomplete(location);
            } catch (error) {
                await this.closeSignIn();
            }
        };
        async autocomplete(value) {
            await this.page.waitForLoadState('load');
            try {
                await this.page.locator(("//div[text()='" + value + "']")).waitFor({ state: 'visible', timeout: 2000 });;
                await this.page.locator(("//div[text()='" + value + "']")).click();
            } catch (error) {
                await this.page.keyboard.press('ArrowDown');
                await this.page.keyboard.press('Enter');
            }
        }
        async dynamicText(value) {
            return ("//div[text()='" + value + "']");
        }
        async selectText(value) {
            await (await this.page.waitForSelector("//*[text()='" + value + "']")).click();
        }
        async checkCurrency(value) {
            let currencyTyp = await this.page.locator(this.currencyBtn).textContent();
            try {
               if (!(currencyTyp.includes(value))) {
                await this.page.waitForLoadState('load');
                await this.page.locator(this.currencyBtn).click();
                const currentFrame = await this.dynamicText(value);
                await this.page.waitForSelector(currentFrame);
                await this.page.locator(currentFrame).click();
            }  
            } catch (error) {
                await this.closeSignIn();
            }
            return await this.page.locator(this.currencyBtn);
        };
        async selectDate(value) {
            let dateValue = value.split('.');
            let day = Number(dateValue[0]);
            let month = await this.monthNo(dateValue[1]);
            let year = Number(dateValue[2]);
            let curDate = new Date();
            let curDay = curDate.getDate();
            let curMonth = curDate.getMonth() + 1;
            let curYear = curDate.getFullYear();
            if (year < curYear || (year == curYear && month < curMonth) || (year == curYear && month == curMonth && day < curDay)) {
                console.log("Invalid Date Selection");
                return null;
            } else {
                let mnthYr;
                await this.page.waitForSelector(this.datePicker);
                month = dateValue[1];
                while (true) {
                    mnthYr = await this.page.locator(this.currentDate).textContent();
                    if (mnthYr.includes(year) && mnthYr.includes(month)) {
                        break;
                    }
                    await this.page.click(this.nextDateButton);
                }
                const date = await this.page.$$(this.dates);
                await this.page.waitForSelector(this.dates + '[1]');
                for (var dt of date) {
                    if ((await dt.textContent()).includes(day)) {
                        await dt.click();
                        await this.page.waitForLoadState('load');
                        return await dt.textContent() + ' ' + mnthYr;
                    }
                }
            }
        }
        async occupancy(type, value) {
            await this.page.waitForSelector(this.occupancyButton);
            await this.page.click(this.occupancyButton);
            switch (type) {
                case 'Individuals':
                    var ir = 2;
                    var irV = 14;
                    break;
                case 'Rooms':
                    var ir = 6;
                    var irV = 20;
                    break;
                default:
                    break;
            }
            var val = Number(value);
            var quantity = Number(await this.page.locator(this.occupancyValue + '[' + irV + ']').textContent());
            while (quantity != val) {
                var i = ir;
                if (val < quantity) {
                    i = (ir - 1);
                }
                await this.page.waitForSelector(this.cntrlButton + '[' + i + ']');
                await this.page.locator(this.cntrlButton + '[' + i + ']').click();
                quantity = Number(await this.page.locator(this.occupancyValue + '[' + irV + ']').textContent());
            }
            await this.selectText('Done');
        }
        async filters(value) {
            var valarray = [value];
            await this.page.waitForLoadState('load');
            await this.closeSignIn();
            var drpdwns = await this.page.$$(this.showDrpdwn);
            for (var drpdwn of drpdwns) {
                if (await drpdwn.isVisible()) {
                    await drpdwn.click();
                }
            }
            var i = 0
            if (value.includes(',')) {
                valarray = value.split(',');
            }
            while (i != valarray.length) {
                let filter = await this.page.locator("(//div[@data-testid='filters-group-container']//*[contains(text(),'" + valarray[i] + "')])[1]");
                let count = await filter.count();
                if (count != 0) {
                    if (await filter.isVisible()) {
                        await filter.click();
                    } else {
                        continue;
                    }
                }
                i++;
            }
            await this.page.waitForSelector(this.availablityBtn);
            //Verifying the applied filters
            valarray = [];
            if (await this.page.locator(this.filterGrouped).isVisible() || await this.page.locator(this.usedFilters1 + '[1]').isVisible()) {
                const appliedFilters = await this.page.locator(this.checkboxes);
                for (let i = 0; i < await appliedFilters.count(); i++) {
                    if (await appliedFilters.nth(i).isChecked()) {
                        valarray.push(await this.page.locator(this.checkboxContent).nth(i).textContent())
                    }
                };
            }
            else {
                console.log("No filters applied");
                valarray.push(null);
            }
            console.log(valarray);
            return valarray;
        }

        async getRules() {
            await this.page.waitForSelector(this.ruleContent);
            const rules = await this.page.$$(this.ruleContent);
            let content = '';
            var i = 0;
            for (const elm of rules) {
                if (i < 5 || (i > 9 && i < 13)) {
                    const val = await elm.textContent();
                    content += "  " + val;
                    if (i % 2 != 0) {
                        content += '\n'
                    }
                }
                i++;
            }
            return content;
        }
        async reservation(member, room) {
            await this.page.waitForLoadState('load');
            let content = '';
            // Get price
            let priceText = await this.page.textContent(this.recommendPrice);
            let price = Number(priceText.replace(/[^0-9.-]+/g, ''))
            // Get tax
            let taxText = await this.page.textContent(this.pricetax);
            let tax = Number(taxText.replace(/[^0-9.-]+/g, ''))
            price += tax;
            await this.page.waitForTimeout(2000);
            var hotelName = await this.page.textContent(this.hotelHeader);
            if (member < 3 && room == 1) {
                await this.page.click(this.reserve);
                await this.page.locator(this.tableRoom).scrollIntoViewIfNeeded();
                var roomName = (await this.page.textContent(this.tableRoom)).trim();
                content = `${hotelName} - ${roomName} for ${member} at ${price}`;
                await this.page.hover(this.cnfrmSubmit);
                await this.page.click(this.cnfrmSubmit);
            } else {
                await this.page.click(this.reserve);
                await this.page.locator(this.tableRoom).scrollIntoViewIfNeeded();
                let recommendContent = await this.page.textContent(this.recommendContent);
                let recommendFor = await this.page.textContent(this.recommendFor);
                content = `${hotelName} - ${recommendContent} for ${recommendFor} at ${price}`;
                try {
                    await this.page.waitForSelector(this.recommendSubmit);
                    await this.page.click(this.recommendSubmit);
                    await this.page.hover(this.cnfrmSubmit);
                    await this.page.click(this.cnfrmSubmit);
                    await this.page.waitForLoadState('load');
                    await this.page.locator(this.firstName).isVisible();
                } catch (error) {
                    await this.page.reload();
                    await this.reservation(member, room);
                }

            }
            return content;
        }
        async fillDetails(firstName, lastName, email, phone, country) {
            await this.page.reload();
            await this.page.waitForLoadState('load');
            await this.page.type(this.firstName, firstName, { delay: 90 });
            await this.page.type(this.lastName, lastName, { delay: 90 });
            await this.page.type(this.email, email, { delay: 90 });
            if (await this.page.locator(this.confirmEmail).isVisible()) {
                await this.page.type(this.confirmEmail, email, { delay: 90 });
            }
            await this.page.selectOption(this.country, country);
            if (await this.page.locator(this.address).isVisible()) {
                await this.page.fill(this.address, 'test');
                await this.page.fill(this.city, 'test');
            }
            await this.page.type(this.phone, '' + phone, { delay: 90 });
            var adres = await this.page.textContent(this.bookadres);
            await this.page.locator(this.finalBtn).scrollIntoViewIfNeeded();
            await this.page.click(this.finalBtn);
            return adres;
        };
        async completeBooking() {
            const arr = [];
            await this.page.waitForLoadState('load');
            await this.closeSignIn();
            await this.page.waitForSelector(this.finalBtn);
            if (await this.page.locator(this.agreeBox).isVisible()) {
                if (!(await this.page.locator(this.agreeBox).isChecked({ timeout: 3000 }))) {
                    await this.page.check(this.agreeBox);
                }
            }
            await this.page.click(this.finalBtn);
            await this.page.waitForLoadState('load');
            await this.page.waitForSelector(this.popUpCls);
            await this.page.click(this.popUpCls);
            arr.push(await this.page.textContent(this.confirmation + '[2]'));
            arr.push(await this.page.textContent(this.confirmation + '[6]'));
            console.log(await this.page.textContent(this.header));
            return arr;
        }
        async filterHotel(budget) {
            const hotelMap = new Map();
            let hotels = await this.page.locator(this.hotels);
            for (var i = 0; i < await hotels.count(); i++) {
                let href, rating, rateCount;
                try {
                    // await hotels.nth(i).locator(this.availableBtn).waitFor({ state: 'visible', timeout: 5000 });
                    await hotels.nth(i).locator(this.availableBtn).scrollIntoViewIfNeeded();
                    href = await hotels.nth(i).locator(this.availableBtn).getAttribute('href');
                    await hotels.nth(i).locator(this.review).isVisible({'timeout':5});
                    let rateText = await hotels.nth(i).locator(this.review).first().textContent();
                    rating = Number(rateText.replace(/[^0-9.-]+/g, ''));
                    let rateCnt = await hotels.nth(i).locator(this.review).last().textContent();
                    rateCount = Number(rateCnt.replace(/[^0-9.-]+/g, ''));
                } catch (error) {
                    continue;
                }
                let priceText = await hotels.locator(this.pricing).nth(i).last().textContent();
                let price = Number(priceText.replace(/[^0-9.-]+/g, ''));
                if (rating >= 7 && rateCount >= 100) {
                    if (price <= budget) {
                        hotelMap.set(await hotels.locator('//h3').nth(i).textContent(),
                            {
                                rating: rating,
                                reviewCount: rateCount,
                                price: price,
                                linkurl:href
                            });
                        console.log(hotelMap.size);
                        if (hotelMap.size > 3) {
                            break;
                        }
                    }
                }
            }
            if (hotelMap.size == 0) {
                console.log("No Hotel in this budget:" + budget);
                budget = budget+1000;
                await this.filterHotel(budget);
            }
            return await this.sortHotelsByScore(hotelMap);
        }
        async sortHotelsByScore(hotelMap) {
            // Convert map to array with score
            const scoredHotels = [];
            for (const [name, data] of hotelMap.entries()) {
                const score = 1000 * (data.rating * Math.log10(1 + data.reviewCount)) / data.price;
                scoredHotels.push({
                    name,
                    ...data,
                    score
                });
            }
            // Sort by score descending
            scoredHotels.sort((a, b) => b.score - a.score);
            // Print each hotel in the sorted array
            scoredHotels.forEach(hotel => {
                console.log(`Hotel: ${hotel.name}, Rating: ${hotel.rating}, Reviews: ${hotel.reviewCount}, Price: ${hotel.price}, Score: ${hotel.score}`);
            });
            return scoredHotels;
        }
    }
