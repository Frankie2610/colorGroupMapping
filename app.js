document.addEventListener("DOMContentLoaded", function () {
    const productInput = document.getElementById("fileInput");
    const processBtn = document.getElementById("processBtn");
    const clearBtn = document.getElementById("clearBtn");
    const meta = document.getElementById("meta");

    let productWorkbook = null;

    /* ================= UTIL ================= */
    function formatSize(bytes) {
        if (!bytes) return "0 KB";
        const k = 1024;
        const sizes = ["Bytes", "KB", "MB", "GB"];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return (bytes / Math.pow(k, i)).toFixed(2) + " " + sizes[i];
    }

    function capitalize(text) {
        return text.charAt(0).toUpperCase() + text.slice(1);
    }

    function extractModelName(title) {
        if (!title) return "UnknownModel";
        const parts = title.trim().split(" ");
        return parts[parts.length - 1] || "UnknownModel";
    }

    function readExcelFile(file, callback) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: "array" });
            callback(wb);
        };
        reader.readAsArrayBuffer(file);
    }

    function showError(message) {
        meta.innerHTML = `<span style="color:#d72c0d;font-weight:600;">❌ ${message}</span>`;
        processBtn.disabled = true;
    }

    /* ================= UPLOAD + PREVIEW ================= */
    productInput.addEventListener("change", function (e) {
        const file = e.target.files[0];
        if (!file) {
            meta.innerText = "Chưa có file. Vui lòng chọn file *.xlsx";
            processBtn.disabled = true;
            return;
        }

        readExcelFile(file, (wb) => {
            productWorkbook = wb;

            const sheet = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            if (!rows.length) {
                showError("File không có dữ liệu.");
                return;
            }

            /* ===== VALIDATE CỘT BẮT BUỘC ===== */
            const headers = Object.keys(rows[0]);

            if (!headers.includes("Product ID")) {
                showError("Thiếu cột bắt buộc: Product ID");
                return;
            }

            /* ===== PREVIEW ===== */
            const totalRows = rows.length;
            const skuRows = rows.filter(
                r => (r["Variant SKU"] || r["SKU"] || "").toString().trim() !== ""
            );

            const skuCount = skuRows.length;

            /* ===== ĐẾM GROUP SẼ TẠO (CHUẨN LOGIC TẠO FILE) ===== */
            /* ===== ĐẾM GROUP THEO CỘT GROUP NAME ===== */
            const previewGroups = {};

            skuRows.forEach(row => {
                const sku = (row["Variant SKU"] || row["SKU"] || "").toString().trim();
                const productId = (row["Product ID"] || "").toString().trim();
                const vendor = (row["Vendor"] || "").toString().trim().toUpperCase();
                const strapColor =
                    (row["Màu dây (product.metafields.custom.m_u_d_y)"] || "").toString().trim();

                // Điều kiện giống hệt lúc tạo Group Name
                if (!sku || !productId || !strapColor) return;

                let prefixLength = 5;
                if (["VERSACE", "FERRAGAMO"].includes(vendor)) prefixLength = 4;
                else if (["MISSONI", "GUESS"].includes(vendor)) prefixLength = 6;
                else if (vendor === "TED BAKER") prefixLength = 7;
                else if (["ADIDAS", "LOCMAN"].includes(vendor)) prefixLength = 8;
                else if (vendor === "FURLA") prefixLength = 10;

                const groupId = sku.substring(0, prefixLength);
                if (!previewGroups[groupId]) {
                    previewGroups[groupId] = {
                        groupId,
                        skuSet: new Set()
                    };
                }

                previewGroups[groupId].skuSet.add(sku);
            });

            /* ===== CHỈ ĐẾM GROUP THỰC SỰ CÓ GROUP NAME ===== */
            const groupNameCount = Object.values(previewGroups)
                .filter(g => g.skuSet.size >= 2)   // chỉ khi đủ điều kiện tạo group
                .length;

            meta.innerHTML = `
                <strong>File đã tải lên:</strong><br>
                📄 <b>${file.name}</b><br>
                📦 Dung lượng: ${formatSize(file.size)}<br>
                🧩 Định dạng: ${file.name.split(".").pop().toUpperCase()}<br><br>

                <strong>Preview dữ liệu:</strong><br>
                🏷️ Tổng số SKU: <b>${totalRows - 1}</b><br>
                🧩 Số group sẽ được tạo: <b>${groupNameCount}</b>
            `;

            processBtn.disabled = groupNameCount === 0;
        });
    });

    /* ================= PROCESS ================= */
    processBtn.addEventListener("click", function () {
        if (!productWorkbook) {
            alert("Bạn chưa upload file sản phẩm!");
            return;
        }

        const sheet = productWorkbook.Sheets[productWorkbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);
        const groups = {};

        rows.forEach(row => {
            const sku = (row["Variant SKU"] || row["SKU"] || "").toString().trim();
            const productId = (row["Product ID"] || "").toString().trim();
            if (!sku || !productId) return;

            const vendor = (row["Vendor"] || "").toString().trim().toUpperCase();

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

            const strapColor =
                (row["Màu dây (product.metafields.custom.m_u_d_y)"] || "").toString().trim();
            if (!strapColor) return;

            if (!groups[skuPrefix]) {
                groups[skuPrefix] = {
                    groupId: skuPrefix,
                    optionId,
                    groupName,
                    values: []
                };
            }

            groups[skuPrefix].values.push({
                productId,
                color: capitalize(strapColor),
                sku
            });
        });

        const output = [];

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
        const exportType =
            document.querySelector('input[name="exportType"]:checked')?.value || "xlsx";

        if (exportType === "csv") {
            const ws = XLSX.utils.json_to_sheet(output);
            const csv = XLSX.utils.sheet_to_csv(ws);
            const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
            const url = URL.createObjectURL(blob);

            const a = document.createElement("a");
            a.href = url;
            a.download = "Group_Mapping_Generated.csv";
            a.click();
            URL.revokeObjectURL(url);
        } else {
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(output);
            XLSX.utils.book_append_sheet(wb, ws, "Group_Mapping");
            XLSX.writeFile(wb, "Group_Mapping_Generated.xlsx");
        }

        meta.innerText = "✅ Tạo file thành công!";
    });

    /* ================= CLEAR ================= */
    clearBtn.addEventListener("click", function () {
        productInput.value = "";
        productWorkbook = null;
        meta.innerText = "Chưa có file. Vui lòng chọn file *.xlsx";
        processBtn.disabled = true;
    });

    processBtn.disabled = true;
});
