using ExcelDataReader.Core;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader;

/// <summary>
/// A range for cells using 0 index positions. 
/// </summary>
/// <param name="FromColumn">The column the range starts in.</param>
/// <param name="FromRow">The row the range starts in.</param>
/// <param name="ToColumn">The column the range ends in.</param>
/// <param name="ToRow">The row the range ends in.</param>
public sealed record CellRange(int FromColumn, int FromRow, int ToColumn, int ToRow)
{
    /// <inheritsdoc/>
    public override string ToString() => $"{FromRow}, {ToRow}, {FromColumn}, {ToColumn}";

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    internal static CellRange Parse(string range)
    {
        int index = range.IndexOf(':');
        if (index < 0 && ReferenceHelper.ParseReference(range, out int column, out int row))
            return new(column - 1, row - 1, column - 1, row - 1);

        if (index >= 0 && range.IndexOf(':', index + 1) < 0)
        {
            ReadOnlySpan<char> span = range;
            if (!ReferenceHelper.ParseReference(span[..index], out int fromColumn, out int fromRow) ||
                !ReferenceHelper.ParseReference(span[(index + 1)..], out int toColumn, out int toRow))
                throw new ExcelReaderException("Invalid cell range reference.");

            // 0 indexed vs 1 indexed
            return new(fromColumn - 1, fromRow - 1, toColumn - 1, toRow - 1);
        }

        throw new ExcelReaderException("Invalid cell range reference.");
    }
#else
    internal static CellRange Parse(string range)
    {
        var fromTo = range.Split(':');
        if (fromTo.Length == 1 && ReferenceHelper.ParseReference(range, out int column, out int row))
            return new(column - 1, row - 1, column - 1, row - 1);

        if (fromTo.Length == 2)
        {
            if (!ReferenceHelper.ParseReference(fromTo[0], out int fromColumn, out int fromRow) ||
                !ReferenceHelper.ParseReference(fromTo[1], out int toColumn, out int toRow))
                throw new ExcelReaderException("Invalid cell range reference.");

            // 0 indexed vs 1 indexed
            return new(fromColumn - 1, fromRow - 1, toColumn - 1, toRow - 1);
        }

        throw new ExcelReaderException("Invalid cell range reference.");
    }
#endif
}