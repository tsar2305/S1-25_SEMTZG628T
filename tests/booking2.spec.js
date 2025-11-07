const { test, expect, chromium } = require('@playwright/test');
const { BookingPage } = require('../pageObject/booking.po');
const { EmailPage } = require('../pageObject/email.po');
var arry = [];
var hotels = [];
// test.setTimeout(60000);
var page1;
test('Search Filtering by date, location@run', async () => {
  // test.slow();
  const browser = await chromium.launch();
  const context = await browser.newContext();
  const page = await context.newPage();

  /**Home Page========================================== */
  const booking = new BookingPage(page);
  //Navigation to Booking Url
  await booking.navigateToPage(booking.bookingUrl);
  await page.waitForSelector(booking.logo);
  //Assuring the page has loaded properly
  expect(await page.locator(booking.logo).isVisible()).toBeTruthy();
  expect(await page.locator(booking.currencyBtn).isVisible()).toBeTruthy();
  //Verifying and Changing Currency
  expect((await booking.checkCurrency('INR'))).toHaveText('INR');
  //Destination Selection and Verification
  await booking.selectDestination('Kolkata')
  expect(await page.locator(booking.destination).getAttribute('value')).toContain('Kolkata');
  //Scheduling Visit via Calender
  await booking.selectDate('2.Dec.2025')
  await booking.selectDate('3.Dec.2025')
  //Occupancy Selection
  await booking.occupancy("Individuals", 2);
  expect(await page.locator(booking.occupancyButton).textContent()).toContain(2 + " adult");
  await booking.occupancy("Rooms", 1);
  expect(await page.locator(booking.occupancyButton).textContent()).toContain(1 + " room");
  await booking.selectText('Search');
  /**Search Result Page========================================== */
  await page.waitForLoadState('load');
  await page.waitForSelector(booking.searchResultHeader);
  var Filter = "Parking";
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
  await page.click(booking.Nopaymnt);
  // await page.click(booking.NoCredCrd);
  await page.waitForSelector(booking.availablityBtn);
  // hotels=(await booking.filterHotel(6000));
  const promise = context.waitForEvent('page');
  await page.click(booking.availablityBtn);
  page1 = await promise;
  await page.close();
})
test("Filtered page handling@confirm", async () => {
  const page = page1;
  // await page.waitForLoadState('load');
  const booking = new BookingPage(page);
  // for (var hotel of hotels) {
  //   try {
  //     await booking.navigateToPage(await hotel.linkurl);
      await page.waitForSelector(booking.reserve);
       await page.waitForTimeout(3000);
      console.log(await booking.getRules());
      await page.waitForTimeout(3000);
      //need change
      arry.push(await booking.reservation(2, 1));
      await page.waitForLoadState('load');
      arry.push(await booking.fillDetails("Rahul", "Maty", 'testing2401@mailinator.com', '915342232', 'India'));
      console.log(arry);
      arry.push(await booking.completeBooking());
      await page.close();
    //   break;
    // } catch (error) {
    //   console.error(`Error processing hotel ${hotel.name}:`, error);
    //   await page.goto('about:blank');
    // }
  // }
})
test('Email Verification@mail', async ({ page }) => {
  // test.slow();
  /**Email Credential Page========================================== */
  const email = new EmailPage(page);
  //Navition to email
  await email.navigateToPage(email.emailUrl);
  //Inputing Credentials to login
  await page.waitForSelector(email.username);
  await page.type(email.username, 'testing2401@mailinator.com', { delay: 120 });
  await page.keyboard.press('Enter');
  await page.waitForLoadState('load');
  /**Inbox Page========================================== */
  const arr = await email.getInbox();
  let array =arry.flat();
  expect(array[2]).toContain(arr[0]);
  expect(array[3]).toBe(arr[1]);
  expect(array[0]).toContain(arr[2]);
  expect(array[1]).toBe(arr[3]);
})