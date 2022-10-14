package com.example.testexcel

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Toast
import io.github.evanrupert.excelkt.*
import kotlinx.android.synthetic.main.activity_main.*
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

class MainActivity : AppCompatActivity() {

    data class Customer(
        val id: String,
        val name: String,
        val address: String,
        val age: Int
    )

    fun findCustomers(): List<Customer> = listOf(
        Customer("1", "Robert", "New York", 32),
        Customer("2", "Bobby", "Florida", 12)
    )

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        btnExport.setOnClickListener { generateXLSXFile() }

    }

    private fun generateXLSXFile() {
        val wb = workbook {
            sheet {
                row {
                    cell("Hello World!")
                }
            }
            sheet("Customers") {
                customersHeader()
                var sum=0.0
                for (customer in findCustomers()){
                    row {
                        cell(customer.id)
                        cell(customer.name)
                        cell(customer.address)
                        cell(customer.age)
                    }
                    sum+=customer.age
                }

                row{
                    cell("UKUPNO:")
                    cell("")
                    cell("")
                    cell(sum)
                }
            }
        }

        SaveXLSXFile(wb)
    }

    private fun SaveXLSXFile(wb: Workbook) {
        val root =
            Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS).toString()
        val myDir = File(root)
        if (!myDir.exists()) {
            myDir.mkdirs()

        }
        val fileName = "Test" + System.currentTimeMillis() + ".xlsx"
        val file = File(myDir, fileName)
        val path = "$root/$fileName"

        var outputStream: FileOutputStream? = null
        try {

            file.createNewFile()
            outputStream = FileOutputStream(file)
            wb.write(path)
            outputStream.flush()
            Toast.makeText(this, "Saved to Downloads", Toast.LENGTH_SHORT).show()


        } catch (t: IOException) {

            Toast.makeText(this, "Not OK", Toast.LENGTH_SHORT).show()
            Log.i("Excel", "Not OK")
            try {
                outputStream?.close()
            } catch (e: Exception) {
                Log.i("Excel", e.stackTraceToString())
            }
        }
    }

    fun Sheet.customersHeader() {
        val headings = listOf("Id", "Name", "Address", "Age")

        val headingStyle = createCellStyle {
            setFont(createFont {
                fontName = "IMPACT"
                color = IndexedColors.PINK.index
            })

            fillPattern = FillPatternType.SOLID_FOREGROUND
            fillForegroundColor = IndexedColors.AQUA.index
        }

        row(headingStyle) {
            headings.forEach { cell(it) }
        }
    }
}