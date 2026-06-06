let selectedFiles = []

const fileInput = document.getElementById("file-input")
const fileList = document.getElementById("file-list")
// ファイル選択時
fileInput.addEventListener("change", function (event) {

    selectedFiles = [...event.target.files]

    renderFileList()
})
// drag & drop
const dropZone = document.getElementById("drop-zone")

dropZone.addEventListener("dragenter", (e) => {
    e.preventDefault()
    dropZone.classList.add("dragover")
})

dropZone.addEventListener("dragover", (e) => {
    e.preventDefault()
    dropZone.classList.add("dragover")
})

dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover")
})

dropZone.addEventListener("drop", (e) => {
    e.preventDefault()
    dropZone.classList.remove("dragover")

    const files = Array.from(
        e.dataTransfer.files
    )

    // 新しい配列を作る

    selectedFiles = [
        ...selectedFiles,
        ...files
    ]

    fileInput.value = ""
    console.log(selectedFiles)

    renderFileList()
})
// 一覧表示機能
function renderFileList() {
    fileList.innerHTML = ""

    selectedFiles.forEach((file, index) => {
        const li = document.createElement("li")
        li.className = "file-item"
        li.dataset.index = index

        const span = document.createElement("span")
        span.textContent = "≡ " + file.name

        // 削除ボタン
        const deleteBtn = document.createElement("button")
        deleteBtn.textContent = "x"
        deleteBtn.className = "delete-btn"
        deleteBtn.onclick = () => {
            selectedFiles.splice(index, 1)
            renderFileList()
        }
        li.appendChild(span)
        li.appendChild(deleteBtn)
        fileList.appendChild(li)
    })
}
// sortable
new Sortable(fileList, {
    animation: 150,
    onEnd: function () {
        const items = fileList.querySelectorAll(".file-item")
        const reordered = []
        items.forEach(item => {
            reordered.push(
                selectedFiles[item.dataset.index]
            )
        })

        selectedFiles = reordered

        renderFileList()
    }
})

document.getElementById("upload-btn")
    .addEventListener("click", async () => {

        if (selectedFiles.length === 0) {
            alert("ファイルを選択してください")
            return
        }

        // ローディング表示
        document.getElementById("loading")
            .style.display = "block"

        const formData = new FormData()

        selectedFiles.forEach(file => {
            formData.append("files", file)
        })
        // ファイル名追加

        const outputName = document.getElementById("output-name").value

        formData.append(
            "output_name",
            outputName
        )

        try {
            const response = await fetch("/upload", {
                method: "POST",
                body: formData
            })
            const blob = await response.blob()

            //デフォルト
            let filename = "自動生成.pptx"

            // Flaskのdownload_name取得
            const disposition = response.headers.get("Content-Disposition")

            console.log(disposition)

            // filename*=utf-8``xxxx 対応
            if (disposition) {
                const utf8Match = disposition.match(/filename\*=UTF-8''(.+)/)

                if (utf8Match && utf8Match[1]) {
                    filename = decodeURIComponent(
                        utf8Match[1]
                    )
                }
            }

            const url = window.URL.createObjectURL(blob)
            const a = document.createElement("a")
            a.href = url

            a.download = filename

            document.body.appendChild(a)

            a.click()

            a.remove()

            window.URL.revokeObjectURL(url)

        } catch (error) {
            alert('生成中にエラーが発生しました')
            console.error(error)
        } finally {
            // ローディング終了
            document.getElementById("loading")
                .style.display = "none"
        }
    })
