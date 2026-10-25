
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const { chromium } = require("playwright");

async function generatePortfolioPdf() {
  const root = process.cwd();
  const htmlFile = path.join(root, "index.html");
  const outputFile = path.join(root, "portfolio.pdf");

  if (!fs.existsSync(htmlFile)) {
    throw new Error("index.html was not found in the repository root.");
  }

  const browser = await chromium.launch({ headless: true });

  try {
    const page = await browser.newPage({
      viewport: { width: 1440, height: 1000 },
      deviceScaleFactor: 1,
    });

    page.on("pageerror", (error) => {
      console.error("Portfolio JavaScript error:", error.message);
    });

    await page.goto(pathToFileURL(htmlFile).href, {
      waitUntil: "load",
      timeout: 60000,
    });

    // Allow page scripts and visual effects to initialize.
    await page.evaluate(async () => {
      await document.fonts.ready;

      await Promise.all(
        Array.from(document.images).map((image) => {
          if (image.complete) return Promise.resolve();
          return new Promise((resolve) => {
            image.addEventListener("load", resolve, { once: true });
            image.addEventListener("error", resolve, { once: true });
          });
        })
      );
    });

    // Disable animations and ensure scroll-reveal content is visible.
    await page.addStyleTag({
      content: `
        *, *::before, *::after {
          animation: none !important;
          transition: none !important;
          scroll-behavior: auto !important;
        }

        .reveal {
          opacity: 1 !important;
          visibility: visible !important;
          transform: none !important;
        }
      `,
    });

    await page.pdf({
      path: outputFile,
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      displayHeaderFooter: false,
      tagged: true,
    });

    const stats = fs.statSync(outputFile);

    if (stats.size < 1000) {
      throw new Error("The generated PDF appears to be empty.");
    }

    console.log(
      `Portfolio PDF generated successfully: ${outputFile} (${stats.size} bytes)`
    );
  } finally {
    await browser.close();
  }
}

generatePortfolioPdf().catch((error) => {
  console.error("PDF generation failed:", error);
  process.exitCode = 1;
});
