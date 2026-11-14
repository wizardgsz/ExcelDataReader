using System.Text;

namespace ExcelDataReader.Core;

internal struct TextAccumulator
{
    private string? _first;
    private StringBuilder? _builder;

    public void Append(string fragment)
    {
        if (fragment.Length == 0)
            return;

        if (_first == null)
            _first = fragment;
        else
            (_builder ??= new StringBuilder(_first)).Append(fragment);
    }

    public override readonly string ToString() => _builder?.ToString() ?? _first ?? string.Empty;
}
