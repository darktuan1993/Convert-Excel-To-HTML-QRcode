const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Đường dẫn tới file Excel gốc
const inputFilePath = path.join(__dirname, 'data.xlsx');

const workbook = XLSX.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
console.log(sheet);
const rows = XLSX.utils.sheet_to_json(sheet);

console.log(rows);



function createHtmlFile(row, index) {
    let html = `
    <!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <style>
    @font-face {
        font-family: "Gothams";
        src: url("https://frontend.mmtsystem.info/font/SVN-Gotham-Ultra.otf") format("opentype");
        font-weight: 900;
        font-style: normal;
    }

    @font-face {
        font-family: "Gothams";
        src: url("https://frontend.mmtsystem.info/font/SVN-Gotham-Light.otf") format("opentype");
        font-weight: 200;
        font-style: normal;
    }

    @font-face {
        font-family: "Gothams";
        src: url("https://frontend.mmtsystem.info/font/SVN-Gotham-Bold.otf") format("opentype");
        font-weight: 600;
        font-style: normal;
    }

    @font-face {
        font-family: "Gothams";
        src: url("https://frontend.mmtsystem.info/font/SVN-Gotham-Regular.otf") format("opentype");
        font-weight: normal;
        font-style: normal;
    }

    /* gothtam */
    @font-face {
        font-family: 'Helvetica Neue Roman';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Roman.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Roman.woff') format('woff');
        font-weight: normal;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Black';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Black.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Black.woff') format('woff');
        font-weight: 900;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Black Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-BlackItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-BlackItalic.woff') format('woff');
        font-weight: 900;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue Heavy';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Heavy.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Heavy.woff') format('woff');
        font-weight: 900;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Heavy Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-HeavyItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-HeavyItalic.woff') format('woff');
        font-weight: 900;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/HelveticaNeue.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/HelveticaNeue.woff') format('woff');
        font-weight: normal;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Italic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Italic.woff') format('woff');
        font-weight: normal;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Bold.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Bold.woff') format('woff');
        font-weight: bold;
        font-style: normal;
    }


    @font-face {
        font-family: 'Helvetica Neue Bold Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-BoldItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-BoldItalic.woff') format('woff');
        font-weight: bold;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue Medium';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Medium.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Medium.woff') format('woff');
        font-weight: 500;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Medium Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-MediumItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-MediumItalic.woff') format('woff');
        font-weight: 500;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue Thin';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Thin.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Thin.woff') format('woff');
        font-weight: 100;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Thin Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-ThinItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-ThinItalic.woff') format('woff');
        font-weight: 100;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue Ultralight';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Ultralight.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Ultralight.woff') format('woff');
        font-weight: 200;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Ultralight Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/HelveticaNeue-UltraLightItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/HelveticaNeue-UltraLightItalic.woff') format('woff');
        font-weight: 100;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue Light';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Light.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-Light.woff') format('woff');
        font-weight: 300;
        font-style: normal;
    }

    @font-face {
        font-family: 'Helvetica Neue Light Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-LightItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-LightItalic.woff') format('woff');
        font-weight: 300;
        font-style: italic;
    }

    @font-face {
        font-family: 'Helvetica Neue Light Italic';
        src: url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-LightItalic.woff2') format('woff2'),
            url('https://rsolution.vn/wp-content/themes/saasland-child/fonts/iCielHelveticaNeue-LightItalic.woff') format('woff');
        font-weight: 300;
        font-style: italic;
    }

    body {
        font-family: Helvetica Neue !important;
    }
</style>
</head>
    `;
    html += `
<body>
<table data-module="header"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/05/fM2VLFkAThoRrsnzIxY6BpUK/_all-in-one/thumbnails/header.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation">
    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs o_pt-lg o_xs-pt-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: #f5f9fd; padding-left: 8px; padding-right: 8px; padding-top: 32px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <div class="ui-resizable-handle ui-resizable-s" style="z-index: 90;"></div>
                <table class="o_block ui-resizable" width="100%" cellspacing="0" cellpadding="0" border="0"
                    role="presentation" style="max-width: 600px;margin: 0 auto;">
                    <tbody>
                        <tr>
                            <td class="o_bg-dark o_px o_py-md o_br-t o_sans o_text" align="center"
                                data-bgcolor="Bg Dark" data-size="Text Default" data-min="12" data-max="20"
                                style="font-family: Helvetica Neue, Arial, sans-serif; margin-top: 0px; margin-bottom: 0px; font-size: 15px; line-height: 27.2px; background-color: #1f3b7f; border-radius: 10px 0px 0px 0px; padding: 24px 16px 20px 16px;">
                                <p style="margin-top: 0px;margin-bottom: 0px;"><a class="o_text-white"
                                        href="https://thefelixthuanan.vn/" data-color="White"
                                        style="text-decoration: none; outline: none; color: rgb(255, 255, 255);">
                                        <!--logo-->
                                        <img src="https://thefelixthuanan.vn/wp-content/uploads/2024/06/c-holdings.png" width="70" height="36"
                                            alt="SimpleApp"
                                            style="max-width: 136px;-ms-interpolation-mode: bicubic;vertical-align: middle;border: 0;line-height: 100%;height: auto;outline: none;text-decoration: none;"
                                            data-crop="false"></a></p>
                            </td>
							<td class="o_bg-dark o_px o_py-md o_br-t o_sans o_text" align="center"
                                data-bgcolor="Bg Dark" data-size="Text Default" data-min="12" data-max="20"
                                style="font-family: Helvetica Neue, Arial, sans-serif; margin-top: 0px; margin-bottom: 0px; font-size: 15px; line-height: 27.2px; background-color: #1f3b7f; padding: 24px 16px 20px;">
                                <p style="margin-top: 0px;margin-bottom: 0px;"><a class="o_text-white"
                                        href="https://thefelixthuanan.vn/" data-color="White"
                                        style="text-decoration: none; outline: none; color: rgb(255, 255, 255);">
                                        <!--logo-->
                                        <img src="https://thefelixthuanan.vn/wp-content/uploads/2024/06/The-Felix_-logo-trang.png" width="136" height="36"
                                            alt="SimpleApp"
                                            style="max-width: 136px;-ms-interpolation-mode: bicubic;vertical-align: middle;border: 0;line-height: 100%;height: auto;outline: none;text-decoration: none;"
                                            data-crop="false"></a></p>
                            </td>
							<td class="o_bg-dark o_px o_py-md o_br-t o_sans o_text" align="center"
                                data-bgcolor="Bg Dark" data-size="Text Default" data-min="12" data-max="20"
                                style="font-family: Helvetica Neue, Arial, sans-serif; margin-top: 0px; margin-bottom: 0px; font-size: 15px; line-height: 27.2px; background-color: #1f3b7f; border-radius: 0px 10px 0px 0px; padding: 24px 16px 20px;">
                                <p style="margin-top: 0px;margin-bottom: 0px;"><a class="o_text-white"
                                        href="https://thefelixthuanan.vn/" data-color="White"
                                        style="text-decoration: none; outline: none; color: rgb(255, 255, 255);">
                                        <!--logo-->
                                        <img src="https://thefelixthuanan.vn/wp-content/uploads/2024/06/LOGO-DKRS-W.png" width="136" height="36"
                                            alt="SimpleApp"
                                            style="max-width: 136px;-ms-interpolation-mode: bicubic;vertical-align: middle;border: 0;line-height: 100%;height: auto;outline: none;text-decoration: none;"
                                            data-crop="false"></a></p>
                            </td>
                        </tr>
                    </tbody>

                </table>
                <!--[if mso]></td></tr></table><![endif]-->
            </td>
        </tr>
    </tbody>
</table>
<table data-module="content"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/09/kZMGx9jluPUJEwbNa7DKCo85/_all-in-one/thumbnails/content.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation">

    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: rgb(245, 249, 253); padding-left: 8px; padding-right: 8px;">

                <table class="o_block" width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation"
                    style="max-width: 600px;margin: 0 auto;">

                    <tbody>
                        <tr>
                            <td class="o_bg-white o_px-md o_py o_sans o_text o_text-secondary" align="center"
                                data-bgcolor="Bg White" data-color="Secondary" data-size="Text Default" data-min="12"
                                data-max="20"
                                style="font-family: Helvetica neue, Arial, sans-serif; margin-top: 0px; margin-bottom: 0px; font-size: 15px; line-height: 24px; background-color: rgb(255, 255, 255); color: rgb(38, 47, 61); padding:30px 16px 24px;">
                                <h2 class="o_heading o_text-dark o_mb-xxs" data-color="Dark" data-size="Heading 2"
                                    data-min="20" data-max="40"
                                    style="font-family: Helvetica Neue, Arial, sans-serif;font-weight: bold;margin-top: 0px;margin-bottom: 4px;color: rgb(36, 43, 61);font-size: 18px;line-height: 29px;">
									<!--THAY ĐỔI TÊN THEO SALE-->
                                    Chào, ${row.Tên}<br></h2>
								<p class="o_heading o_text-dark o_mb-xxs" data-color="Dark" data-size="Heading 2"
                                    data-min="20" data-max="40"
                                    style="font-family: Helvetica Neue, Arial, sans-serif;font-weight: bold;margin-top: 0px;margin-bottom: 4px;color: rgb(36, 43, 61);font-size: 14px;line-height: 29px;">
                                    Dưới đây là mã QR-CODE check-in tham dự "Chuỗi đào tạo chuyên sâu"</br> dự án The Felix của bạn!<br></p>
                            </td>
                        </tr>

                    </tbody>
                </table>

            </td>
        </tr>

    </tbody>
</table>
<table data-module="image-full" width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation">

    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: rgb(245, 249, 253); padding-left: 8px; padding-right: 8px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <table class="o_block" width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation"
                    style="max-width: 600px;margin: 0 auto;">

                    <tbody>
                        <tr>
                            <td class="o_bg-white o_sans o_text o_text-secondary" align="center" data-bgcolor="Bg White"
                                data-size="Text Default" data-min="12" data-max="20" data-color="Secondary"
                                style="font-family: Helvetica, Arial, sans-serif;margin-top: 0px;margin-bottom: 0px;font-size: 16px;line-height: 24px;background-color: #ffffff;color: #424651;">
								<!--THAY ĐỔI QR THEO SALE-->
                                <p style="margin-top: 0px;margin-bottom: 0px;"><img class="o_img-full"
                                        src="${row.Mã}" width="400" alt=""
                                        style="max-width: 400px;-ms-interpolation-mode: bicubic;vertical-align: middle;border: 0;line-height: 100%;height: auto;outline: none;text-decoration: none;width: 100%;"
                                        data-crop="false"></p>
                            </td>
                        </tr>

                    </tbody>
                </table>
                <!--[if mso]></td></tr></table><![endif]-->
            </td>
        </tr>

    </tbody>
</table>
<table data-module="subtitle"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/09/kZMGx9jluPUJEwbNa7DKCo85/_all-in-one/thumbnails/subtitle.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation" style="opacity: 1;">

    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: rgb(245, 249, 253); padding-left: 8px; padding-right: 8px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <table class="o_block" width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation"
                    style="max-width: 600px;margin: 0 auto;">

                    <tbody>
                        <tr>
                            <td class="o_bg-white o_px-md o_pt" align="center" data-bgcolor="Bg White"
                                style="background-color: #ffffff;padding-left: 24px;padding-right: 24px;padding-top: 30px;">
                                <h4 class="o_heading o_text-dark" data-color="Dark" data-size="Heading 4" data-min="10"
                                    data-max="26"
                                    style="font-family: Helvetica neue, Arial, sans-serif;font-weight: bold;margin-top: 0px;margin-bottom: 0px;color: #242b3d;font-size: 17px;line-height: 23px;">
                                    Mã rút thăm trúng thưởng của bạn tại sự kiện Kick-Off The Felix là:<span style="    color: #e63757;">
									<!--THAY ĐỔI CODE RÚT THĂM THEO SALE-->
									${row.ID}</span>
                                </h4>

                            </td>
                        </tr>

                    </tbody>
                </table>
                <!--[if mso]></td></tr></table><![endif]-->
            </td>
        </tr>

    </tbody>
</table>
<table data-module="content"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/09/kZMGx9jluPUJEwbNa7DKCo85/_all-in-one/thumbnails/content.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation">

    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: rgb(245, 249, 253); padding-left: 8px; padding-right: 8px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <table class="o_block" width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation"
                    style="max-width: 600px;margin: 0 auto;">

                    <tbody>
                        <tr>
                            <td class="o_bg-white o_px-md o_py o_sans o_text o_text-secondary" align="center"
                                data-bgcolor="Bg White" data-color="Secondary" data-size="Text Default" data-min="12"
                                data-max="20"
                                style="font-family: Helvetica neue, Arial, sans-serif;margin-top: 0px;margin-bottom: 0px;font-size: 14px;line-height: 24px;background-color: #ffffff;color: #424651;padding-left: 50px;padding-right: 50px;padding-top: 16px;padding-bottom: 16px;">
                                <p style="margin-top: 0px;margin-bottom: 0px;"><strong>Mã QR-CODE</strong> được sử dụng để check-in khi tham dự <strong>"Chuỗi đào tạo chuyên sâu"</strong> dự án The Felix. Khi tham dự bạn sẽ có cơ hội nhận những phần thưởng hấp dẫn thông qua <strong>"Mã rút thăm trúng thưởng"</strong> tại sự kiện Kick-Off The Felix 😉🍾🥂
                                </p>
                                <p></p>
                            </td>
                        </tr>

                    </tbody>
                </table>
                <!--[if mso]></td></tr></table><![endif]-->
            </td>
        </tr>

    </tbody>
</table>
<table data-module="content"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/06/DRTuKOVNXe4Qg5mywzx3S20d/_all-in-one/thumbnails/content.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation">

    <tbody>
        <tr>

            <td class="o_bg-light o_px-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: #f5f9fd;padding-left: 8px;padding-right: 8px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <table class="o_block" width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation"
                    style="max-width: 600px;margin: 0 auto;">

                    <tbody>
                        <tr>

                            <td class="o_bg-white o_px-md o_py o_sans o_text o_text-secondary" align="center"
                                data-bgcolor="Bg White" data-color="Secondary" data-size="Text Default" data-min="12"
                                data-max="20"
                                style="font-family: Helvetica Neue, Arial, sans-serif;margin-top: 0px;margin-bottom: 0px;font-size: 15px;line-height: 24px;background-color: #ffffff;color: #424651;padding-left: 24px;padding-right: 24px;padding-top: 16px;padding-bottom: 16px;">
                                <table align="center" cellspacing="0" cellpadding="0" border="0" role="presentation">
                                    <tbody>
                                        <tr>
                                            <td width="250" class="o_bb-light"
                                                style="font-size: 8px;line-height: 8px;height: 8px;border-bottom: 1px solid #9aa1a5;"
                                                data-border-bottom-color="Border Light">&nbsp; </td>
                                        </tr>
                                    </tbody>
                                </table>
                                <p style="margin-top: 20px;margin-bottom: 0px;">Cám ơn bạn đã xem,</p>
                                <div style="font-weight: 600;">THE FELIX | KHẮC HỌA HẠNH PHÚC - ĐIỂM TÔ TƯƠNG LAI</div>
                                <p></p>
                            </td>
                        </tr>

                    </tbody>
                </table>
                <!--[if mso]></td></tr></table><![endif]-->
            </td>
        </tr>

    </tbody>
</table>
<table data-module="footer-2cols"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/05/fM2VLFkAThoRrsnzIxY6BpUK/_all-in-one/thumbnails/footer-2cols.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation" class="">
    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs o_pb-lg o_xs-pb-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: #f5f9fd;padding-left: 8px;padding-right: 8px;padding-bottom: 0px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <div class="ui-resizable-handle ui-resizable-s" style="z-index: 90;"></div>
                <table class="o_block ui-resizable" width="100%" cellspacing="0" cellpadding="0" border="0"
                    role="presentation" style="max-width: 600px;margin: 0 auto;">
                    <tbody>
                        <tr>
                            <td class="o_re o_bg-dark o_px o_pb-lg o_br-b" align="center" data-bgcolor="Bg Dark"
                                style="font-size: 0;vertical-align: top;background-color: #1f3b7f;border-radius: 0px 0px 10px 10px;padding-left: 16px;padding-right: 16px;padding-bottom: 15px;">
                                <!--[if mso]><table cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td width="200" align="left" valign="top" style="padding:0px 8px;"><![endif]-->
                                <div class="o_col o_col-4"
                                    style="display: inline-block;vertical-align: top;width: 100%;max-width: 400px;">
                                    <div style="font-size: 32px;line-height: 32px;height: 20px;" class="">&nbsp;
                                    </div>
                                    <div class="o_px-xs o_sans o_text-xs o_text-dark_light o_left o_xs-center"
                                        data-color="Dark Light" data-size="Text XS" data-min="10" data-max="18"
                                        style="font-family: Helvetica Neue, Arial, sans-serif;margin-top: 0px;margin-bottom: 0px;font-size: 14px;line-height: 21px;color: rgb(245, 249, 253);text-align: center;padding-left: 8px;padding-right: 8px;">

                                        <p class="o_mb-xs" style="margin-top: 0px;margin-bottom: 8px;">
                                            hotro@dkrs.com.vn | 1900 2663
                                        </p>

                                    </div>
                                </div>
                            </td>
                        </tr>
                    </tbody>

                </table>
                <!--[if mso]></td></tr></table><![endif]-->

            </td>
        </tr>
    </tbody>
</table>
<table data-module="footer-2cols"
    data-thumb="http://www.stampready.net/dashboard/editor/user_uploads/zip_uploads/2019/12/05/fM2VLFkAThoRrsnzIxY6BpUK/_all-in-one/thumbnails/footer-2cols.png"
    width="100%" cellspacing="0" cellpadding="0" border="0" role="presentation" class="">
    <tbody>
        <tr>
            <td class="o_bg-light o_px-xs o_pb-lg o_xs-pb-xs" align="center" data-bgcolor="Bg Light"
                style="background-color: #f5f9fd; padding-left: 8px; padding-right: 8px; padding-bottom: 32px;">
                <!--[if mso]><table width="632" cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td><![endif]-->
                <div class="ui-resizable-handle ui-resizable-s" style="z-index: 90;"></div>
                <table class="o_block ui-resizable" width="100%" cellspacing="0" cellpadding="0" border="0"
                    role="presentation" style="max-width: 600px;margin: 0 auto;">
                    <tbody>
                        <tr>
                            <td class="o_re o_bg-dark o_px o_pb-lg o_br-b" align="center" data-bgcolor="Bg Dark"
                                style="font-size: 0;vertical-align: top;background-color: #242b3d00;border-radius: 0px 0px 10px 10px;padding-left: 16px;padding-right: 16px;padding-bottom: 15px;">
                                <!--[if mso]><table cellspacing="0" cellpadding="0" border="0" role="presentation"><tbody><tr><td width="200" align="left" valign="top" style="padding:0px 8px;"><![endif]-->
                                <div class="o_col o_col-4"
                                    style="display: inline-block;vertical-align: top;width: 100%;max-width: 400px;">
                                    <div style="font-size: 32px;line-height: 32px;height: 10px;" class="">&nbsp;
                                    </div>
                                    <div class="o_px-xs o_sans o_text-xs o_text-dark_light o_left o_xs-center"
                                        data-color="Dark Light" data-size="Text XS" data-min="10" data-max="18"
                                        style="font-family: Helvetica Neue, Arial, sans-serif;margin-top: 0px;margin-bottom: 0px;font-size: 12px;line-height: 21px;color: rgb(218, 218, 218);text-align: center;padding-left: 8px;padding-right: 8px;">
                                        <p class="o_mb-xs" style="margin-top: 0px;margin-bottom: 0px;">17 Mê Linh, Phường 19, Quận Bình Thạnh, Tp. Hồ Chí Minh</p>

                                        <p style="margin-top: 0px;margin-bottom: 0px;">
                                            <a class="o_text-dark_light o_underline"
                                                href="https://dkrs.com.vn/gioi-thieu/" data-color="Dark Light"
                                                style="text-decoration: underline;outline: none;color: rgb(218, 218, 218);">
                                                Về chúng tôi
                                            </a>
                                            <span class="o_hide-xs">&nbsp; • &nbsp;</span><br class="o_hide-lg"
                                                style="display: none;font-size: 0;max-height: 0;width: 0;line-height: 0;overflow: hidden;mso-hide: all;visibility: hidden;">
                                            <a class="o_text-dark_light o_underline"
                                                href="https://dkrs.com.vn/lien-he/"
                                                data-color="Dark Light"
                                                style="text-decoration: underline;outline: none;color: rgb(218, 218, 218);">Liên hệ</a> <span class="o_hide-xs">&nbsp; • &nbsp;</span><br
                                                class="o_hide-lg"
                                                style="display: none;font-size: 0;max-height: 0;width: 0;line-height: 0;overflow: hidden;mso-hide: all;visibility: hidden;">
                                            <a class="o_text-dark_light o_underline"
                                                href="https://dkrs.com.vn/danh-sach-tuyen-dung/"
                                                data-color="Dark Light"
                                                style="text-decoration: underline;outline: none;color: rgb(218, 218, 218);">Tuyển dụng</a>
                                        </p>
                                        <p class="o_mb-xs" style="margin-top: 0px;margin-bottom: 0px; font-size: 11px;">
                                            <a style="color:rgb(218, 218, 218)" href="#">(Ngưng nhận email từ chúng
                                                tôi.)</a></p>

                                    </div>
                                </div>
                                <!--[if mso]></td><td width="400" align="right" valign="top" style="padding:0px 8px;"><![endif]-->

                                <!--[if mso]></td></tr></table><![endif]-->
                            </td>
                        </tr>
                    </tbody>

                </table>
                <!--[if mso]></td></tr></table><![endif]-->
                <div class="o_hide-xs" style="font-size: 64px; line-height: 64px; height: 64px;">&nbsp; </div>
            </td>
        </tr>
    </tbody>
</table>
</body>


    
    `;

    // html += `<h2>${row.Tên}</h2>`;

    // Object.keys(row).forEach(header => {
    //     console.log(row.Tên);
    //     html += `<h2>${row.Tên}</h2>`;
    // });

    // Tạo tên file mới dựa trên chỉ số dòng
    const outputFilePath = path.join(__dirname, `row-${index + 1}.html`);
    fs.writeFileSync(outputFilePath, html);

    console.log(`Created: ${outputFilePath}`);
}




rows.forEach((row, index) => {
    createHtmlFile(row, index);
});
