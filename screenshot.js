const { chromium } = require('playwright');

(async () => {
    const browser = await chromium.launch();
    const page = await browser.newPage({
        viewport: { width: 1440, height: 900 }
    });

    await page.goto('http://localhost:8080');
    await page.waitForLoadState('networkidle');

    // Wait a moment for fonts to load
    await page.waitForTimeout(2000);

    // Take hero screenshot
    await page.screenshot({
        path: '/home/claude/mikevertx/preview-hero.png',
        fullPage: false
    });

    // Scroll to ladder section and screenshot
    await page.evaluate(() => {
        document.querySelector('#ladder').scrollIntoView();
    });
    await page.waitForTimeout(500);
    await page.screenshot({
        path: '/home/claude/mikevertx/preview-ladder.png',
        fullPage: false
    });

    // Scroll to clients section
    await page.evaluate(() => {
        document.querySelector('#clients').scrollIntoView();
    });
    await page.waitForTimeout(500);
    await page.screenshot({
        path: '/home/claude/mikevertx/preview-clients.png',
        fullPage: false
    });

    // Scroll to portfolio
    await page.evaluate(() => {
        document.querySelector('#portfolio').scrollIntoView();
    });
    await page.waitForTimeout(500);
    await page.screenshot({
        path: '/home/claude/mikevertx/preview-portfolio.png',
        fullPage: false
    });

    // Full page screenshot
    await page.screenshot({
        path: '/home/claude/mikevertx/preview-full.png',
        fullPage: true
    });

    await browser.close();
    console.log('Screenshots captured!');
})();
