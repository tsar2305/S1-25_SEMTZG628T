const { test, expect, chromium } = require('@playwright/test');
const { BookingPage } = require('../pageObject/booking.po');
const { EmailPage } = require('../pageObject/email.po');
const { Data } = require('../pageObject/data.po');
const { faker } = require('@faker-js/faker');
const { xlsx } = require('xlsx');
var browser, page, context;
var page1 = null;
let testData, hotels, arry = [];

test.describe('DataDriven Hotel Booking', () => {

  const data = new Data();
  //reading testData Sheet
  testData = data.readExcel('data/testData.xlsx', 'Sheet1');
  // array of dataset
  console.log('Loaded testData:', testData);

  test.beforeEach('Open Url', async () => {
    //Making new Page 
    browser = await chromium.launch();
    context = await browser.newContext();
    if (page1 == null) {
      page = await context.newPage();
    } else {
      page = page1;
    }
    page.waitForLoadState('load');
  });

  test.afterEach('Close all Pages', async () => {
    //Closing Page after Execution
    await page.close();
    await data.writeExcel('data/testData.xlsx', 'Sheet1', testData);
  })
  test('Fill Excel with random data@random', async () => {
    const booking = new BookingPage(page);
    console.log('Filling Excel with Random Data');
      for (const dt of testData) {
        const randomData =await booking.ranData();
        data.updateExcel(testData, dt.SerialNo, 'FirstName',  randomData[0]);
        data.updateExcel(testData, dt.SerialNo, 'LastName',  randomData[1]);
        data.updateExcel(testData, dt.SerialNo, 'Email',  randomData[2]);
        data.updateExcel(testData, dt.SerialNo, 'Phone',  randomData[3]);
        data.updateExcel(testData, dt.SerialNo, 'Country',  randomData[4]);
        data.updateExcel(testData, dt.SerialNo, 'Filter',  randomData[5]);
        data.updateExcel(testData, dt.SerialNo, 'Currency',  randomData[6]);
        data.updateExcel(testData, dt.SerialNo, 'Member',  randomData[7]);
        data.updateExcel(testData, dt.SerialNo, 'Room',  randomData[8]);
        data.updateExcel(testData, dt.SerialNo, 'Budget',  randomData[9]);
        data.updateExcel(testData, dt.SerialNo, 'Arrival',  randomData[10]);
        data.updateExcel(testData, dt.SerialNo, 'Departure',  randomData[11]);
        data.updateExcel(testData, dt.SerialNo, 'Destination',  randomData[12]);
      }
  })
  for (const dt of testData) {
    test('Search Filtering by date, location ' + dt.SerialNo + '@run', async () => {
      // test.slow();
      /**Home Page========================================== */
      const booking = new BookingPage(page);
      //Navigation to Booking Url
      await booking.navigateToPage(booking.bookingUrl);
      await page.waitForSelector(booking.logo);
      //Assuring the page has loaded properly
      expect(await page.locator(booking.logo).isVisible()).toBeTruthy();
      expect(await page.locator(booking.currencyBtn).isVisible()).toBeTruthy();
      //Verifying and Changing Currency
      expect((await booking.checkCurrency(dt.Currency))).toHaveText(dt.Currency);
      //Destination Selection and Verification
      await booking.selectDestination(dt.Destination)
      expect.soft(await page.locator(booking.destination).getAttribute('value')).toContain(dt.Destination);
      //Scheduling Visit via Calender
      let configDate = (dt.Arrival).split('.');
      configDate = Number(configDate[0]) + ' ' + configDate[1];
      expect(await booking.selectDate(dt.Arrival)).toContain(configDate);
      configDate = (dt.Departure).split('.');
      configDate = Number(configDate[0]) + ' ' + configDate[1];
      expect(await booking.selectDate(dt.Departure)).toContain(configDate);
      //Occupancy Selection
      await booking.occupancy("Individuals", dt.Member);
      expect(await page.locator(booking.occupancyButton).textContent()).toContain(dt.Member + " adult");
      await booking.occupancy("Rooms", dt.Room);
      expect(await page.locator(booking.occupancyButton).textContent()).toContain(dt.Room + " room");
      await booking.selectText('Search');
      /**Search Result Page========================================== */
      await page.waitForSelector(booking.searchResultHeader);
      var Filter = dt.Filter;
      if (Filter != null) {
        let filters = await booking.filters(Filter);
        for (var filter of filters) {
          if (await filter != null) {
            if (Filter.includes(',')) {
              const str1Parts = Filter.split(',');
              expect(str1Parts.some(part => filter.includes(part))).toBeTruthy();
            }
            else {
              expect(filter.includes(Filter)).toBeTruthy();
            }
          }
        }
      }
      await page.waitForSelector(booking.Nopaymnt);
      await page.click(booking.Nopaymnt);
      // await page.click(booking.NoCredCrd);      
      /**Filter Hotels by review ratings and Price */
      await page.waitForSelector(booking.availablityBtn);
      try {
        hotels = await booking.filterHotel(dt.Budget);
        await page.click(booking.availablityBtn);
      } catch (error) {
        await page.click(booking.availablityBtn);
        hotels = [{linkurl:await this.page.getUrl()}];
      }
      /**Open new Window Promise*/
      const promise = context.waitForEvent('page');
      if (promise != null) {
        page1 = await promise;
        await page1.waitForLoadState('load')
      }
    })
    test("Filtered page handling " + dt.SerialNo + '@run', async () => {
      const booking = new BookingPage(page);
      for (var hotel of hotels) {
        try {
          /**Preferred Hotel Url*/
          await booking.navigateToPage(hotel.linkurl);
          await page.waitForSelector(booking.reserve);
          /**Rules update on data sheet */
          await data.updateExcel(testData, dt.SerialNo, 'Rules', await booking.getRules());
          await page.waitForTimeout(3000);
          /**Details and Address of Booking update on Data Sheet*/
          await data.updateExcel(testData, dt.SerialNo, 'Details', await booking.reservation(dt.Member, dt.Room));
          await page.waitForLoadState('load');
          await data.updateExcel(testData, dt.SerialNo, 'Address', await booking.fillDetails(dt.FirstName, dt.LastName, dt.Email, dt.Phone, dt.Country));
          /**Complete Booking */
          const arr = await booking.completeBooking()
          /**Update PIN and PNR on Data Sheet */
          await data.updateExcel(testData, dt.SerialNo, 'PNR', arr[0]);
          await data.updateExcel(testData, dt.SerialNo, 'PIN', arr[1]);
          break;
        } catch (error) {
          console.error(`Error processing hotel ${hotel.name}:`, error);
          /**Next Preferred Option on Error in previous flow */
          await page.goto('about:blank');
        }
      }
      page1 = null;
    })
    test('Email Verification ' + dt.SerialNo + '@run', async () => {
      test.slow();
      /**Email Credential Page========================================== */
      const email = new EmailPage(page);
      //Navition to email
      await email.navigateToPage(email.emailUrl);
      //Inputing Credentials to login
      await page.waitForSelector(email.username);
      await page.type(email.username, dt.Email, { delay: 120 });
      await page.keyboard.press('Enter');
      await page.waitForLoadState('load');
      /**Inbox Page========================================== */
      arry = await email.getInbox();
      /**Validate the Contents generated from Booking Site and Extracted from Email Inbox*/
      expect(arry[0]).toBe(dt.PNR);
      expect(arry[1]).toBe(dt.PIN);
      expect(dt.Details).toContain(arry[2]);
      expect(arry[3]).toBe(dt.Address);
      if(test.fail()==true){
        await data.updateExcel(testData, dt.SerialNo, 'Status', 'Fail');
      }else{
        await data.updateExcel(testData, dt.SerialNo, 'Status', 'Pass');
      }
    })
  }
});