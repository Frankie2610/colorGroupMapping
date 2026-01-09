const exportType =
    document.querySelector('input[name="exportType"]:checked')?.value || "xlsx";
document.addEventListener("DOMContentLoaded", function () {
    const productInput = document.getElementById("fileInput");
    const processBtn = document.getElementById("processBtn");
    const clearBtn = document.getElementById("clearBtn");
    const meta = document.getElementById("meta");

    let productWorkbook = null;

    function readExcelFile(file, callback) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: "array" });
            callback(wb);
        };
        reader.readAsArrayBuffer(file);
    }

    productInput.addEventListener("change", function (e) {
        const file = e.target.files[0];
        if (!file) {
            meta.innerText = "Chưa có file. Vui lòng chọn file *.xlsx";
            return;
        }
        readExcelFile(file, (wb) => {
            productWorkbook = wb;
            meta.innerText = "Đã tải: " + file.name;
            console.log("Product file loaded OK");
        });
    });

    function capitalize(text) {
        return text.charAt(0).toUpperCase() + text.slice(1);
    }

    function extractModelName(title) {
        if (!title) return "UnknownModel";
        const parts = title.trim().split(" ");
        return parts[parts.length - 1] || "UnknownModel";
    }


    processBtn.addEventListener("click", function () {
        if (!productWorkbook) {
            alert("Bạn chưa upload file sản phẩm!");
            return;
        }

        const sheet = productWorkbook.Sheets[productWorkbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);
        const groups = {};

        // Gom nhóm theo prefix SKU
        rows.forEach(row => {
            const sku = (row["Variant SKU"] || row["SKU"] || "").toString().trim();
            const productId = (row["Product ID"] || "").toString().trim();
            if (!sku || !productId) return;

            const vendor = (row["Vendor"] || "").toString().trim().toUpperCase();

            // Xác định số ký tự prefix theo Vendor
            let prefixLength = 5;
            if (["VERSACE", "FERRAGAMO"].includes(vendor)) prefixLength = 4;
            else if (["PHILIPP PLEIN", "VERSUS BY VERSACE"].includes(vendor)) prefixLength = 5;
            else if (["MISSONI", "GUESS"].includes(vendor)) prefixLength = 6;
            else if (vendor === "TED BAKER") prefixLength = 7;
            else if (vendor === "ADIDAS") prefixLength = 8;
            else if (vendor === "LOCMAN") prefixLength = 8;
            else if (vendor === "FURLA") prefixLength = 10;

            const skuPrefix = sku.substring(0, prefixLength);
            const optionId = skuPrefix;
            const modelName = extractModelName(row["Title"]);
            const groupName = `${vendor}-${skuPrefix}-${modelName}`;

            // ✅ LẤY MÀU DÂY (STRING)
            const strapColor =
                (row["Màu dây (product.metafields.custom.m_u_d_y)"] || "").toString().trim();

            if (!strapColor) return; // không có màu → bỏ

            if (!groups[skuPrefix]) {
                groups[skuPrefix] = {
                    groupId: skuPrefix,
                    optionId: optionId,
                    groupName: groupName,
                    values: []
                };
            }

            groups[skuPrefix].values.push({
                productId: productId,
                color: capitalize(strapColor),
                modelName: modelName,
                sku: sku
            });
        });

        const output = [];

        // Tạo output, bỏ qua nhóm có <2 mẫu
        Object.keys(groups).forEach(prefix => {
            const g = groups[prefix];
            // ❗ Không đủ 2 sản phẩm thì KHÔNG tạo group
            if (g.values.length < 2) return;

            // Dòng 1
            output.push({
                "Group ID": g.groupId,
                "Group Name": g.groupName,
                "Product ID": "",
                "Combination ID": "",
                "Option ID": "",
                "Option Name": "",
                "Style On Page": "",
                "Style On Card": "",
                "Value ID": "",
                "Value Name": "",
                "Swatch Style": "",
                "Swatch Color 1": "",
                "Swatch Color 2": "",
                "Swatch Image": ""
            });

            // Dòng 2
            output.push({
                "Group ID": g.groupId,
                "Group Name": "",
                "Product ID": "",
                "Combination ID": "",
                "Option ID": g.optionId,
                "Option Name": "Màu sắc",
                "Style On Page": "Image Swatch With Price",
                "Style On Card": "Circle Swatch",
                "Value ID": "",
                "Value Name": "",
                "Swatch Style": "",
                "Swatch Color 1": "",
                "Swatch Color 2": "",
                "Swatch Image": ""
            });

            // Dòng 3+ (các giá trị)
            g.values.forEach(v => {
                output.push({
                    "Group ID": g.groupId,
                    "Group Name": "",
                    "Product ID": v.productId,
                    "Combination ID": v.sku,
                    "Option ID": g.optionId,
                    "Option Name": "",
                    "Style On Page": "",
                    "Style On Card": "",
                    "Value ID": v.sku,
                    "Value Name": v.color,
                    "Swatch Style": "First Image",
                    "Swatch Color 1": "",
                    "Swatch Color 2": "",
                    "Swatch Image": ""
                });
            });
        });

        // Xuất file
        // Xuất file theo định dạng được chọn
        // Xuất file theo định dạng được chọn (RADIO)
        const exportType =
            document.querySelector('input[name="exportType"]:checked')?.value || "xlsx";

        if (exportType === "csv") {
            const ws = XLSX.utils.json_to_sheet(output, { skipHeader: false });
            const csv = XLSX.utils.sheet_to_csv(ws);

            const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
            const url = URL.createObjectURL(blob);

            const a = document.createElement("a");
            a.href = url;
            a.download = "Group_Mapping_Generated.csv";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

        } else {
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(output, { skipHeader: false });
            XLSX.utils.book_append_sheet(wb, ws, "Group_Mapping");
            XLSX.writeFile(wb, "Group_Mapping_Generated.xlsx");
        }


        meta.innerText = "Tạo file thành công!";
    });

    clearBtn.addEventListener("click", function () {
        productInput.value = "";
        productWorkbook = null;
        meta.innerText = "Chưa có file. Vui lòng chọn file *.xlsx";
    });
});
