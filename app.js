const exportTypeSelect = document.getElementById("exportType");
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

    function extractColor(bodyHtml) {
        if (!bodyHtml) return "Chưa có";
        const lowered = bodyHtml.toLowerCase();
        const COLOR_MAP = [
            "đen", "nâu", "xám", "xanh", "đỏ", "vàng", "vàng hồng",
            "bạc", "trắng", "hồng", "tím", "cam", "be", "navi", "navy", "trắng có", "cam", "xanh rêu", "xanh lá"
        ];
        for (let c of COLOR_MAP) {
            if (lowered.includes(c)) return capitalize(c);
        }
        return "Chưa có";
    }

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

            // Xác định số ký tự để lấy làm prefix theo Vendor
            let prefixLength = 5; // mặc định
            debugger;
            if (["VERSACE", "FERRAGAMO"].includes(vendor)) prefixLength = 4;
            else if (["PHILIPP PLEIN", "VERSUS BY VERSACE"].includes(vendor)) prefixLength = 5;
            else if (["TED BAKER", "MISSONI", "GUESS"].includes(vendor)) prefixLength = 6;
            else if (vendor === "FURLA") prefixLength = 10;
            else if (vendor === "LOCMAN") prefixLength = 8;

            const skuPrefix = sku.substring(0, prefixLength);
            const optionId = skuPrefix;
            const modelName = extractModelName(row["Title"]);
            const groupName = `${vendor}-${skuPrefix}-${modelName}`;
            const color = extractColor(row["Body (HTML)"]);

            if (!groups[skuPrefix]) {
                groups[skuPrefix] = {
                    groupId: skuPrefix,
                    optionId: optionId,
                    groupName: groupName,
                    values: [],
                    models: new Set()
                };
            }

            groups[skuPrefix].values.push({
                productId: productId,
                color: color,
                modelName: modelName,
                sku: sku
            });
            groups[skuPrefix].models.add(modelName);
        });



        const output = [];

        // Tạo output, bỏ qua nhóm có <2 mẫu
        Object.keys(groups).forEach(prefix => {
            const g = groups[prefix];
            // if (g.models.size < 2) return; // bỏ qua nhóm chỉ có 1 mẫu

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
        const exportType = exportTypeSelect?.value || "xlsx";

        if (exportType === "csv") {
            // CSV
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
            // XLSX (mặc định)
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
