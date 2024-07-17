#include "ThumbnailProvider.h"
#include <Shlwapi.h>
#include <memory>
#include <wingdi.h>
#include "WebpReader.h"
#include <poppler/cpp/poppler-page.h>
#include <poppler/cpp/poppler-document.h>
#include <cairo.h>
#include <memory>
#include <fstream>

namespace fastpdfext
{
    using namespace std;

    ThumbnailProvider::ThumbnailProvider() :
        ref_count_{1},
        stream_{nullptr}
    {
    }

    ThumbnailProvider::~ThumbnailProvider()
    {
        if (stream_)
            stream_->Release();
    }

#pragma region IUnknown

    HRESULT ThumbnailProvider::QueryInterface(const IID& riid, void** ppvObject)
    {
        static const QITAB qit[]{
            QITABENT(ThumbnailProvider, IInitializeWithStream),
            QITABENT(ThumbnailProvider, IThumbnailProvider),
            {0}
        };
        return QISearch(this, qit, riid, ppvObject);
    }

    ULONG ThumbnailProvider::AddRef()
    {
        return InterlockedIncrement(&ref_count_);
    }

    ULONG ThumbnailProvider::Release()
    {
        const long ref = InterlockedDecrement(&ref_count_);
        if (ref == 0)
            delete this;
        return ref;
    }

#pragma endregion

#pragma region IThumbnailProvider

    HRESULT ThumbnailProvider::GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha)
    {
        STATSTG stat;
        HRESULT hr = stream_->Stat(&stat, STATFLAG::STATFLAG_NONAME);
        if (!SUCCEEDED(hr))
            return hr;

        const unique_ptr<uint8_t[]> data = make_unique<uint8_t[]>(stat.cbSize.QuadPart);
        ULONG bytes_read;
        hr = stream_->Read(data.get(), static_cast<ULONG>(stat.cbSize.QuadPart), &bytes_read);
        if (!SUCCEEDED(hr))
            return hr;

        // 检查文件类型，如果是 PDF 文件则处理
        std::string file_signature(reinterpret_cast<char*>(data.get()), 4);
        if (file_signature == "%PDF") {
            // 使用 Poppler 库处理 PDF 文件
            std::string pdf_path = "temp.pdf";
            std::ofstream pdf_file(pdf_path, std::ios::binary);
            pdf_file.write(reinterpret_cast<char*>(data.get()), bytes_read);
            pdf_file.close();

            auto document = poppler::document::load_from_file(pdf_path);
            if (!document) {
                return E_FAIL;
            }

            auto page = document->create_page(0);
            if (!page) {
                return E_FAIL;
            }

            double width, height;
            page->size(&width, &height);

            double scale = static_cast<double>(cx) / std::max(width, height);
            int scaled_width = static_cast<int>(width * scale);
            int scaled_height = static_cast<int>(height * scale);

            BITMAPINFO bmi;
            bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
            bmi.bmiHeader.biHeight = -scaled_height;
            bmi.bmiHeader.biWidth = scaled_width;
            bmi.bmiHeader.biPlanes = 1;
            bmi.bmiHeader.biBitCount = 32;
            bmi.bmiHeader.biCompression = BI_RGB;
            *pdwAlpha = WTSAT_ARGB;

            TBYTE* bytes;
            HBITMAP bmp = CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS,
                reinterpret_cast<void**>(&bytes),
                nullptr, 0);

            if (!bmp) {
                return E_OUTOFMEMORY;
            }

            cairo_surface_t* surface = cairo_image_surface_create_for_data(
                bytes, CAIRO_FORMAT_ARGB32, scaled_width, scaled_height, scaled_width * 4);

            if (cairo_surface_status(surface) != CAIRO_STATUS_SUCCESS) {
                DeleteObject(bmp);
                return E_FAIL;
            }

            cairo_t* cr = cairo_create(surface);
            if (cairo_status(cr) != CAIRO_STATUS_SUCCESS) {
                cairo_surface_destroy(surface);
                DeleteObject(bmp);
                return E_FAIL;
            }

            cairo_scale(cr, scale, scale);
            page->render(cr);

            // 在这里添加红色标记
            int markWidth = 10; // 红色标记的宽度
            int markHeight = 10; // 红色标记的高度

            for (int y = 0; y < markHeight; ++y) {
                for (int x = 0; x < markWidth; ++x) {
                    int pixelIndex = (y * scaled_width + x) * 4; // 每个像素4个字节（RGBA）
                    bytes[pixelIndex] = 0;     // 蓝色通道
                    bytes[pixelIndex + 1] = 0; // 绿色通道
                    bytes[pixelIndex + 2] = 255; // 红色通道
                    bytes[pixelIndex + 3] = 255; // Alpha通道
                }
            }

            cairo_destroy(cr);
            cairo_surface_destroy(surface);

            *phbmp = bmp;
            return S_OK;
        }

        // 处理 WebP 文件的现有代码保持不变
        const WebpReader reader;
        INT webp_width, webp_height;
        BOOLEAN webp_alpha;
        hr = reader.ReadWebpHeader(data.get(), bytes_read, &webp_width, &webp_height, &webp_alpha);

        INT scaled_width, scaled_height;
        CalcScaledBmpSize(webp_width, webp_height, cx, &scaled_width, &scaled_height);

        if (!SUCCEEDED(hr))
            return hr;

        BITMAPINFO bmi;
        bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
        bmi.bmiHeader.biHeight = -scaled_height;
        bmi.bmiHeader.biWidth = scaled_width;
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;
        *pdwAlpha = webp_alpha ? WTSAT_ARGB : WTSAT_RGB;

        TBYTE* bytes;
        HBITMAP bmp = CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS,
            reinterpret_cast<void**>(&bytes),
            nullptr, 0);

        hr = bmp ? S_OK : E_OUTOFMEMORY;
        if (!SUCCEEDED(hr))
            return hr;

        WebpReadInfo ri{
            data.get(),
            bytes_read,
            scaled_width,
            scaled_height,
            bytes
        };

        hr = reader.ReadAsBitmap(&ri);
        if (SUCCEEDED(hr))
        {
            // 在这里添加红色标记
            int markWidth = 10; // 红色标记的宽度
            int markHeight = 10; // 红色标记的高度

            for (int y = 0; y < markHeight; ++y) {
                for (int x = 0; x < markWidth; ++x) {
                    int pixelIndex = (y * scaled_width + x) * 4; // 每个像素4个字节（RGBA）
                    bytes[pixelIndex] = 0;     // 蓝色通道
                    bytes[pixelIndex + 1] = 0; // 绿色通道
                    bytes[pixelIndex + 2] = 255; // 红色通道
                    bytes[pixelIndex + 3] = 255; // Alpha通道
                }
            }

            *phbmp = bmp;
        }

        return hr;
    }

    void ThumbnailProvider::CalcScaledBmpSize(const INT webp_width, const INT webp_height,
                                              const INT cx, INT* scaled_width, INT* scaled_height) const
    {
        if (webp_width > cx || webp_height > cx)
        {
            if (webp_width > webp_height)
            {
                const double ratio = 1.0 * cx / webp_width;
                *scaled_width = cx;
                *scaled_height = static_cast<uint32_t>(round(ratio * webp_height));
            }
            else if (webp_height > webp_width)
            {
                const double ratio = 1.0 * cx / webp_height;
                *scaled_width = static_cast<uint32_t>(round(ratio * webp_width));
                *scaled_height = cx;
            }
            else
            {
                *scaled_width = cx;
                *scaled_height = cx;
            }
        }
        else
        {
            *scaled_width = webp_width;
            *scaled_height = webp_height;
        }
    }

#pragma endregion

#pragma region IInitializeWithStream

    HRESULT ThumbnailProvider::Initialize(IStream* pstream, DWORD grfMode)
    {
        HRESULT hr = E_UNEXPECTED;
        if (!stream_)
            hr = pstream->QueryInterface(&stream_);
        return hr;
    }

#pragma endregion
}
