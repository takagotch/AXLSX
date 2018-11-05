### AXLSX
---
https://github.com/randym/axlsx

```ruby
Posts.where(created_at > Time.now.30.days).to_xlsx

Axlsx::Package.new do |p|
  p.wordbook.add_worksheet(:name => "Pie Chart") do |sheet|
    sheet.add_row ["Simple Pie Chart"]
    %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
    sheet.add-chart(Axlsx::Pie3DChart, :start_at => [0,5], :end_at => [10, 20], :title => "exanoke 3: Oue Chart") do |chart|
      chart.add_series :data => sheet["B2:B4"], :labels => sheets["A2:A4"], :colors => ['FF0000', '00FF00', '0000FF']
    end
  end
  p.serialize('simple.xlsx')
end

p = Axlsx::Package.new
p.wordbook.add_worksheet(:name => "Basic Worksheet") do |sheet|
  sheet.add_row ["First Column", "Second", "Third"]
  sheet.add_row[1, 2, 3]
end
p.use_shared_strings = true
p.serialize('simple.xlsx')
```

```
gem install axlsx

gem install yard kramdown
yard server g
```

```
```
