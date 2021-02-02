# # vendor = wb[vendors]
# for vendor_name in vendors:
#     vendor = wb[vendor_name]
#
#     # set periode
#     vendor["L8"] = str(today.strftime("%M")).title()
#     vendor["L9"] = today.strftime(f"1-{today.strftime('%d %B %Y')}")
#     print(today.strftime("%B"))
#     print(f"1-{today.strftime('%d %B %Y')}")
#
# insert image
    # skull_img = Image("skull.png")
    # title_img = Image("title.png")
    # one_img = Image("one.png")
    #
    # vendor.add_image(skull_img, "A1")
    # vendor.add_image(title_img, "B2")
    # title_img.width = 700
    # vendor.add_image(one_img, "K1")
#     # set stok awal = stok akhir
#     for stok_akhir in vendor["J"][12:]:
#         if stok_akhir.value:
#             print(f"{vendor.title} stok akhir: {stok_akhir.value}"+f" stok awal: {vendor[f'F{i}'].value}")
#             vendor[f"F{i}"].value = stok_akhir.value
#             i += 1
#         else:
#             break
#
#     # set penjualan = 0
#     for penjualan in vendor["H"][12:]:
#         if penjualan.value:
#             penjualan.value = 0
#
#     # set keterangan = kosong
#     for keterangan in vendor["L"][12:]:
#         keterangan.value = ""
#
# # print(stok_awal)
#
# # for vendor in vendors:
# #     vendor = wb[vendor]
# #
# #     for stok_akhir in vendor["J"][12:15]:
# #         print(f"{vendor.title}: {stok_akhir.value}")
#
# # ade_doclo = wb["ADE DOCLO"]
# # # print(ade_doclo["F"])
# #
# # # ade_doclo["A10"] = "test"
# #
# # # set stok awal = stok akhir periode lalu
# # for stok_akhir in ade_doclo["J"][12:15]:
# #     ade_doclo[f"F{i}"].value = stok_akhir.value
# #     i += 1
# #
# # # set penjualan = 0
# # for penjualan in ade_doclo["H"][12:15]:
# #     penjualan.value = 0
# #
# # # set keterangan kosong
# # for keterangan in ade_doclo["L"][12:15]:
# #     keterangan.value = ""
#